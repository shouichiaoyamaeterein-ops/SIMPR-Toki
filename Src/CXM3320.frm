VERSION 5.00
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{2BB9E22C-2B4F-11D4-806A-00000EA57F9E}#5.0#0"; "CISNFT~1.OCX"
Begin VB.Form CXM3320 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ポイントメンテナンス"
   ClientHeight    =   11235
   ClientLeft      =   2205
   ClientTop       =   3300
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   Begin Cis3D_v60.CIS3D Message 
      Height          =   960
      Left            =   3990
      Top             =   5640
      Visible         =   0   'False
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   1693
      BackColor       =   16761024
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
      cBoderWidth     =   5
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   720
         Left            =   105
         Top             =   120
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   1270
         BackColor       =   16761024
         Caption         =   "【表示中】しばらくお待ち下さい。"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HG丸ｺﾞｼｯｸM-PRO"
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
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin Cis3D_v60.CIS3D B1_Area2 
      Height          =   2310
      Left            =   12765
      Top             =   8175
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   4075
      BackColor       =   8421504
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   2175
         Left            =   75
         Top             =   60
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3836
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
         Begin Cis3D_v60.CIS3D NoImage 
            Height          =   2145
            Left            =   15
            Top             =   15
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   3784
            BackColor       =   14737632
            Caption         =   "No Image"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "HGSｺﾞｼｯｸE"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   14737632
            cBoderColor2    =   14737632
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin VB.Image B_IMG 
            Height          =   2145
            Left            =   15
            Stretch         =   -1  'True
            Top             =   15
            Width           =   2130
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   510
      Top             =   2190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox H1_Area2 
      BackColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   3675
      ScaleHeight     =   795
      ScaleWidth      =   8805
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   555
      Width           =   8865
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   675
         Left            =   60
         Top             =   60
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   1191
         ForeColor       =   192
         Caption         =   "イメージ ファイル パス"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D lb_IMGPath 
            Height          =   315
            Left            =   60
            Top             =   270
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   556
            BackColor       =   16777152
            Caption         =   "C:\Documents and Settings\administrator\My Documents"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cPositionX      =   60
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   7935
      Left            =   1995
      ScaleHeight     =   7875
      ScaleWidth      =   10485
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2550
      Width           =   10545
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   7710
         Left            =   75
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   75
         Width           =   10320
         _Version        =   196608
         _ExtentX        =   18203
         _ExtentY        =   13600
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
         FormatString    =   "    |ﾎﾟｲﾝﾄ|イメージファイル                                                        "
         Rows            =   21
         Cols            =   5
         BackColor       =   12648384
         ForeColorFixed  =   128
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   360
         SelectionMode   =   1
         BackColorAlternate=   12648384
         Begin VB.PictureBox Input_Area 
            Appearance      =   0  'ﾌﾗｯﾄ
            BackColor       =   &H00000000&
            BorderStyle     =   0  'なし
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   585
            ScaleHeight     =   345
            ScaleWidth      =   9435
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   405
            Width           =   9435
            Begin Cis3D_v60.CIS3D Back_Point 
               Height          =   345
               Left            =   0
               Top             =   0
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   609
               BackColor       =   12640511
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
               Begin CisNfText_V60.CisNfText iB1_Point 
                  Height          =   300
                  Left            =   60
                  TabIndex        =   0
                  Top             =   45
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   529
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
            Begin Cis3D_v60.CIS3D PB_Image 
               Height          =   330
               Left            =   9045
               Top             =   0
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   582
               Caption         =   "･･･"
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
               cAlingnment     =   8
               cPositionX      =   30
               cPositionY      =   90
            End
            Begin CisText_V60.CisText iB1_IMG 
               Height          =   360
               Left            =   675
               TabIndex        =   1
               Top             =   0
               Width           =   8520
               _ExtentX        =   15028
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
               Text            =   $"CXM3320.frx":0000
               MaxLength       =   200
               IMEMode         =   0
            End
            Begin VB.TextBox B_TabDummy 
               Height          =   270
               Left            =   8745
               TabIndex        =   2
               Text            =   "Text1"
               Top             =   60
               Width           =   210
            End
         End
      End
   End
   Begin VB.Timer Timer2 
      Left            =   705
      Top             =   420
   End
   Begin VB.Timer Timer1 
      Left            =   285
      Top             =   420
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   45
      TabIndex        =   3
      Text            =   "Dummy"
      Top             =   405
      Width           =   195
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
      Left            =   8130
      ScaleHeight     =   735
      ScaleWidth      =   4350
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1545
      Width           =   4410
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1755
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
         Left            =   2610
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
         Left            =   3465
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   900
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   45
         Top             =   75
         Visible         =   0   'False
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
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 ポイントメンテナンス  】"
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   11535
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
Attribute VB_Name = "CXM3320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   ポイントメンテナンス
'**       フォームID    :   CXM3320
'**       処理概要      :
'**
'**       作  成  日    :   2005/11/08  By CIS SIMPR-A
'**       変  更  日    :
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

    Dim wImageUNC       As String
    Dim mImageUNC       As String
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
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = All
        .FixedHeight = 430
        .PatanMax = 1
        .SelectCol = 3
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
' イメージ格納先
    With CisFun
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_Section = "Option"
        .INI_Key = "Image":       .INI_Default = gIniExe
        .GetIni
        If Right(RTrim(.INI_String), 1) <> "\" Then
            lb_IMGPath = RTrim(.INI_String) & "\"
        Else
            lb_IMGPath = RTrim(.INI_String)
        End If
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2005/12/13 ADD START >
    
        .FE_FullPath = lb_IMGPath
        .GetUNCPath
        lb_IMGPath = RTrim(CisFun.FE_UNCPath)
        mImageUNC = StrConv(CisFun.FE_UNCPath, 1)
    End With

' 初期表示
    Call PB_ENT_Click
    
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
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                  Call PB_ENT_Click
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
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
    If ProcHB = "B1" And SyoriKB = "DEL" Then
        SyoriKB = ""
        ProcHB = "B1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        Call VsGridToText(vsGrid1.Row)
        iB1_Point.SetFocus
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Dummy.Enabled = False
    DoEvents
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
        ProcHB = "B1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        vsGrid1.SetFocus
    Else
        SyoriKB = ""
        Call SyoriName(SyoriKB)
        vsGrid1.Row = 1
        Call VsGridToText(vsGrid1.Row)
        iB1_Point.SetFocus
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iB1_IMG" Then Call PB_Image_Click: Exit Sub
End Sub
'+----------------------------+
'+    ｲﾒｰｼﾞﾌｧｲﾙ指定           +
'+----------------------------+
Private Sub PB_Image_Click()
     On Error Resume Next
     
     CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly
     If RTrim(iB1_IMG) = "" Then
        CommonDialog1.FileName = RTrim(lb_IMGPath) & "Image"
     Else
        CommonDialog1.FileName = RTrim(lb_IMGPath) & iB1_IMG
     End If
     gStr = "JPEG ｲﾒｰｼﾞ  (*.jpg)|*.jpg"
     gStr = gStr & "|すべてのﾋﾟｸﾁｬﾌｧｲﾙ| *.bmp;;*dib;*.gif;*.jpg;*.wmf;*.emf"
     gStr = gStr & "|ﾋﾞｯﾄﾏｯﾌﾟ    (*.bmp;*.dib)|*.bmp;*.dib"
     gStr = gStr & "|GIF ｲﾒｰｼﾞ   (*.gif)|*.gif"
     gStr = gStr & "|ﾒﾀﾌｧｲﾙ      (*.wmf;*.emf)|*.wmf;*.emf"
     gStr = gStr & "|すべてのﾌｧｲﾙ (*.*)|*.*"
     CommonDialog1.Filter = gStr
    ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
     CommonDialog1.CancelError = True
     CommonDialog1.ShowOpen
     If Err = 32755 Then
    '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
    '必要なｿｰｽｺｰﾄﾞを入力してください。
        Exit Sub
     End If
     If Err = 20477 Then
        With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    ファイル名が不正です。            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
        End With
        Exit Sub
     End If
     If Trim(CommonDialog1.FileName) <> "" Then
         iB1_IMG = Replace(CommonDialog1.FileName, lb_IMGPath, "")
     End If
     
     Call ImageDisp
     iB1_IMG.SetFocus
     On Error GoTo 0
End Sub
'+-----------------------------+
'+  iB1_Point
'+-----------------------------+
Private Sub iB1_Point_GotFocus()
    Set ActObj = iB1_Point
    Back_Point.BackColor = gIGotSel
    Back_Point.Tag = "S"
End Sub
Private Sub iB1_Point_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Point_LostFocus()
    vsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 1) = RTrim(iB1_Point)
    Back_Point.BackColor = gILostSel
End Sub
'+-----------------------------+
'+  iB1_IMG
'+-----------------------------+
Private Sub iB1_IMG_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_IMG
    Back_Point.Tag = ""
End Sub
Private Sub iB1_IMG_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_IMG_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    vsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 2) = RTrim(iB1_IMG)
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If Back_Point.Tag = "" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyDown)
        iB1_Point.SetFocus
        Exit Sub
    Else
        If Val(Input_Area.Tag) = 1 Then
           iB1_Point.SetFocus
        Else
           Call TextToVsGrid(Val(Input_Area.Tag))
           Call Grid_RowChange(vbKeyUp)
           iB1_Point.SetFocus
        End If
    End If
End Sub

'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   Dim W_GridColor     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
     Case "DEL"
          SyoriNM = "削除"
          K_Sykbnm.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
     Case "LOK"
          SyoriNM = "照会"
          K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
          PB_DEL.Visible = False
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
   Call CisFrm.TimeSet(1)
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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    H1Mode = False
    H1Color = gPLostSel
    H2Mode = False
    H2Color = gPLostSel
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
       Case "H2"
          H2Mode = True
          H2Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Visible = B1Mode
    
    PB_DEL.Visible = B1Mode
    If SyoriKB = "DEL" Then
        PB_CAN.Visible = True
        Input_Area.Visible = Not B1Mode
    Else
        PB_CAN.Visible = False
        Input_Area.Visible = B1Mode
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
                .Enabled = True
            End If
            If .Name Like "Chk" & wbc_Enm & "_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
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
        If Not BodyDisp Then GoTo ReturnPress_Ed
        vsGrid1.Row = 1
        Call VsGridToText(vsGrid1.Row)
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            Call BodyDisp
            vsGrid1.SetFocus
        Else
            If Not Body1Chk Then GoTo ReturnPress_Ed
            If Not DBPut Then GoTo ReturnPress_Ed
            ProcHB = "B1"
            Call DispChange(ProcHB)
            Call HeadBodyClear("B1")
            Call BodyDisp
            vsGrid1.Row = 1
            Call VsGridToText(vsGrid1.Row)
        End If
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      Body部表示処理
'****************************
Private Function BodyDisp() As Boolean
    BodyDisp = False
    
    Call HeadBodyClear("B")
    Message.Visible = True
' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " SELECT * FROM ポイントマスタ"
    gSL_Select = gSL_Select & " ORDER BY ポイント"
    If POMRead(gSL_Select, 0) Then
        Call GridSet
    End If
    Call POMClose
       
    Message.Visible = False
    BodyDisp = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    

    Call TextToVsGrid(Val(Input_Area.Tag))
    Input_Area.Visible = False

'<< 空白行を詰める >>チェック
    gCnt2 = 0
    With vsGrid1
        For gCnt1 = 1 To .Rows - 1
    
        'ポイント
            If Trim(.TextMatrix(gCnt1, 1)) = "" And Trim(.TextMatrix(gCnt1, 2)) = "" Then
                For gCnt2 = gCnt1 + 1 To .Rows - 1
                    If Trim(.TextMatrix(gCnt2, 1)) <> "" Or Trim(.TextMatrix(gCnt2, 2)) <> "" Then
                        For gCnt3 = 1 To .Cols - 1
                            .TextMatrix(gCnt1, gCnt3) = .TextMatrix(gCnt2, gCnt3)
                            .TextMatrix(gCnt2, gCnt3) = ""
                        Next gCnt3
                        Exit For
                    End If
                Next gCnt2
            End If
        
        Next gCnt1
    End With
'<< 明細 >>チェック
    With vsGrid1
        For gCnt1 = 1 To .Rows - 1
    
        'ポイント
            If Trim(.TextMatrix(gCnt1, 1)) = "" And Trim(.TextMatrix(gCnt1, 2)) <> "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ポイントを入力して下さい            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_Point
                GoTo Body1Chk_Err
            End If
        
        ' ポイント半角チェック
            If RTrim(.TextMatrix(gCnt1, 1)) <> "" And Not HalfSizeChk(.TextMatrix(gCnt1, 1)) Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(2) = "    ポイントコードは半角にて入力して下さい            "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                Set ErrObj = iB1_Point
                GoTo Body1Chk_Err
            End If
        
        'イメージ
            If Trim(.TextMatrix(gCnt1, 1)) <> "" And Trim(.TextMatrix(gCnt1, 2)) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    イメージファイルを入力して下さい            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_IMG
                GoTo Body1Chk_Err
            End If
            If Left(Trim(.TextMatrix(gCnt1, 2)), 1) = "\" Or _
                Left(Trim(.TextMatrix(gCnt1, 2)), 1) = "." Or _
                Left(Trim(.TextMatrix(gCnt1, 2)), 1) = "/" Or _
                InStr(.TextMatrix(gCnt1, 2), ":") <> 0 Then
                With CisFun
                    .FE_FullPath = vsGrid1.TextMatrix(gCnt1, 2)
                    .GetUNCPath
                    wImageUNC = StrConv(RTrim(CisFun.FE_UNCPath), 1)
                End With
                .TextMatrix(gCnt1, 2) = Replace(wImageUNC, mImageUNC, "")
            End If
            If Left(Trim(.TextMatrix(gCnt1, 2)), 1) = "\" Or _
                Left(Trim(.TextMatrix(gCnt1, 2)), 1) = "." Or _
                Left(Trim(.TextMatrix(gCnt1, 2)), 1) = "/" Or _
                InStr(.TextMatrix(gCnt1, 2), ":") <> 0 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(1) = "    イメージファイル　指定異常！            "
                    .MB_MSG(3) = "    \ / .　は不可！            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_IMG
                GoTo Body1Chk_Err
            End If
            
            If RTrim(.TextMatrix(gCnt1, 1)) <> "" Then
        'ポイントの重複確認
                For gCnt2 = gCnt1 + 1 To .Rows - 1
                    If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) Then
                        With CisFun
                            .MB_Lines = 4
                            .MB_MSG(1) = "    　【ポイント　重複】            "
                            .MB_MSG(3) = "    " & gCnt1 & " 行目と " & gCnt2 & " 行目     "
                            .MB_Title = ""
                            .MB_Button = Error
                            .MBOX
                        End With
                        Set ErrObj = iB1_Point
                        GoTo Body1Chk_Err
                    End If
                Next gCnt2
                
        'イメージファイルの存在確認
                gStr = lb_IMGPath & Trim(.TextMatrix(gCnt1, 2))
               If Dir(gStr) = "" Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(1) = "    　【イメージファイルが存在しません】            "
                        .MB_MSG(3) = "    " & gCnt1 & " 行目 " & gStr
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    Set ErrObj = iB1_IMG
                    GoTo Body1Chk_Err
               End If
            End If
        
        Next gCnt1
    End With

    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid1.Row = gCnt1
    Call VsGridToText(vsGrid1.Row)
    Input_Area.Visible = True
    ErrObj.SetFocus
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
'  ( 選択確認 )
    Call CisVsGrid.vsColor(, &HC0C0FF)
    If RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     行を選択してください。      "
            .MB_Title = "行選択"
            .MB_Button = Error
            .MBOX
            vsGrid1.SetFocus
            Exit Function
        End With
    End If
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 5
        .MB_MSG(1) = "     削除処理を行います      "
        .MB_MSG(3) = "    " & vsGrid1.TextMatrix(vsGrid1.Row, 1) & " : " & vsGrid1.TextMatrix(vsGrid1.Row, 2)
        .MB_Title = "削除処理"
        .MB_MSG(5) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call CisVsGrid.vsColor(全部)
            vsGrid1.SetFocus
            Exit Function
        End If
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        ' 発注サイクルマスタ
        .SQL = "DELETE FROM ポイントマスタ "
        .SQL = .SQL & " WHERE ポイント = '" & RTrim(RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 1))) & "'"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

        On Error GoTo 0
    End With
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

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います      "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Input_Area.Visible = True
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    With CisDB
        On Error GoTo DBPut_Err
' ( 削除処理 )
        ' 発注サイクルマスタ
        .SQL = "TRUNCATE TABLE ポイントマスタ "
        .DBExec

        On Error GoTo 0
    End With
'   マスタ項目移送
    With vsGrid1
        For gIndex = 1 To .Rows - 1
            If RTrim(.TextMatrix(gIndex, 1)) = "" Then Exit For
            POM.ポイント = RTrim(.TextMatrix(gIndex, 1))
            POM.イメージ = RTrim(.TextMatrix(gIndex, 2))
            POM.作成者 = gTanto
        ' ( 追加処理 )
            Call POMInsert
        Next gIndex
    End With
    
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
        Call .FixedSet(中中, 右中, あり, 4, 0, "№")
        Call .FixedSet(中中, 左中, あり, 5, 1, "ﾎﾟｲﾝﾄ")
        Call .FixedSet(中中, 左中, あり, 72, 2, "イメージファイル")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 3, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2")
        .InitDisp
    End With
    
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
    With vsGrid1
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
    Do Until Not POM_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid1
            .TextMatrix(ID, GridCol(0)) = ID
            .TextMatrix(ID, GridCol(1)) = RTrim(POM.ポイント)
            .TextMatrix(ID, GridCol(2)) = RTrim(POM.イメージ)
        End With
        Call POMReadNext
    Loop

    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = vsGrid1.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid1.Rows - 1 Then
            If vsGrid1.TextMatrix(gInt, 1) = "" Then
                GoTo Grid_RowChange_End
            Else
                vsGrid1.Rows = vsGrid1.Rows + 1
                vsGrid1.TextMatrix(gInt + 1, 0) = gInt + 1
            End If
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
        If gInt + CisVsGrid.DispMax > vsGrid1.Rows - 1 Then
            gInt = vsGrid1.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    vsGrid1.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(vsGrid1.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid1
        .TextMatrix(tRow, 1) = RTrim(iB1_Point)
        .TextMatrix(tRow, 2) = RTrim(iB1_IMG)
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid1
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, vsGrid1.CellTop, Input_Area.Width, vsGrid1.CellHeight
        'ポイント
        iB1_Point = RTrim(.TextMatrix(tRow, 1))
        'イメージ
        iB1_IMG = RTrim(.TextMatrix(tRow, 2))
        'イメージ表示
        Call ImageDisp
    End With
End Sub
Private Sub ImageDisp()
    'イメージ表示
    If RTrim(iB1_IMG) <> "" And _
        Left(Trim(iB1_IMG), 1) <> "\" And _
        Left(Trim(iB1_IMG), 1) <> "." And _
        Left(Trim(iB1_IMG), 1) <> "/" Then
        If Dir(RTrim(lb_IMGPath) & iB1_IMG) <> "" Then
            Set B_IMG.Picture = LoadPicture(RTrim(lb_IMGPath) & iB1_IMG)
            NoImage.Left = 3000
            DoEvents
        Else
            Set B_IMG.Picture = LoadPicture("")
            NoImage.Left = 15
        End If
    Else
        Set B_IMG.Picture = LoadPicture("")
        NoImage.Left = 3000
    End If
End Sub

Private Sub vsGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
'        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
'    Call TextToVsGrid(Val(Input_Area.Tag))
'    Call VsGridToText(vsGrid1.Row)
End Sub
Private Sub VSGrid1_Click()
    If SyoriKB <> "DEL" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        If vsGrid1.TextMatrix(vsGrid1.Row, 0) = "" Then Exit Sub
        Input_Area.Visible = True
        Call VsGridToText(vsGrid1.Row)
        On Error Resume Next
        If ActObj.Visible Then ActObj.SetFocus
        On Error GoTo 0
    End If
End Sub
Private Sub VSGrid1_DblClick()
    If SyoriKB = "DEL" Then
        Call PB_ENT_Click
    End If
End Sub

Private Sub vsGrid1_RowColChange()
    If SyoriKB = "DEL" Then
        iB1_IMG = vsGrid1.TextMatrix(vsGrid1.Row, 2)
        Call ImageDisp
    End If
End Sub

Private Sub vsGrid1_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub


