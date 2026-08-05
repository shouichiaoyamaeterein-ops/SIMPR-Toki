VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CXU0120 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "売上実績一覧表"
   ClientHeight    =   11190
   ClientLeft      =   1950
   ClientTop       =   1650
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4290
      ScaleHeight     =   285
      ScaleWidth      =   6555
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   6615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1605
      Top             =   4590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5625
      Top             =   7290
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      Caption         =   "終  了"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   4305
      Left            =   4260
      ScaleHeight     =   4418.27
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2565
      Width           =   6660
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   4140
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   7303
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Height          =   570
            Left            =   1965
            TabIndex        =   17
            Top             =   2325
            Width           =   3240
            Begin VB.OptionButton H1Op_List 
               BackColor       =   &H00C0C0C0&
               Caption         =   "品番単位"
               Height          =   255
               Index           =   1
               Left            =   1830
               TabIndex        =   5
               Top             =   210
               Width           =   1350
            End
            Begin VB.OptionButton H1Op_List 
               BackColor       =   &H00C0FFC0&
               Caption         =   "取引先マスタ"
               Height          =   255
               Index           =   0
               Left            =   150
               TabIndex        =   4
               Top             =   210
               Value           =   -1  'True
               Width           =   1455
            End
         End
         Begin VB.ComboBox H1Comb_Koku 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1980
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   3
            Top             =   1995
            Visible         =   0   'False
            Width           =   1605
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Index           =   0
            Left            =   660
            Top             =   1785
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "契約先"
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
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   1980
            TabIndex        =   1
            Top             =   1095
            Width           =   1440
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   75
            TabIndex        =   14
            Top             =   3060
            Width           =   6345
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   4380
               TabIndex        =   8
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   2505
               TabIndex        =   7
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1065
               TabIndex        =   6
               Top             =   270
               Value           =   -1  'True
               Width           =   1260
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D28 
            Height          =   615
            Index           =   1
            Left            =   60
            Top             =   60
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   1085
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
            Begin VB.PictureBox B1_Area 
               BackColor       =   &H00FFC0FF&
               Height          =   555
               Left            =   45
               ScaleHeight     =   495
               ScaleWidth      =   6210
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   30
               Width           =   6270
               Begin Cis3D_v60.CIS3D lb_msg 
                  Height          =   405
                  Left            =   60
                  Top             =   60
                  Width           =   6090
                  _ExtentX        =   10742
                  _ExtentY        =   714
                  BackColor       =   16761087
                  Caption         =   "売上実績一覧表　発行を行います"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   14.25
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cBoderColor1    =   16761087
                  cBoderColor2    =   16761087
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   6
                  cBoderStyle     =   2
               End
            End
         End
         Begin Cis3D_v60.CIS3D lb_PrtDev 
            Height          =   375
            Left            =   75
            Top             =   3690
            Width           =   5970
            _ExtentX        =   10530
            _ExtentY        =   661
            BackColor       =   16777152
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
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D PB_Print 
            Height          =   345
            Left            =   6075
            Top             =   3720
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   609
            Caption         =   "..."
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
            cPositionX      =   30
            cPositionY      =   -90
         End
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   345
            Index           =   1
            Left            =   570
            Top             =   1095
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "実績年月"
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
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   1980
            TabIndex        =   2
            Top             =   1755
            Width           =   1020
            _ExtentX        =   1799
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   360
            Left            =   3345
            Top             =   1755
            Width           =   2535
            _ExtentX        =   4471
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D PB_Torcd 
            Height          =   360
            Left            =   3000
            Top             =   1755
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            ForeColor       =   65280
            Caption         =   "▼"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cButton         =   -1  'True
            cPositionX      =   20
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Index           =   1
            Left            =   660
            Top             =   2025
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "工  区  "
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
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Index           =   2
            Left            =   660
            Top             =   2505
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "ﾘｽﾄ形式"
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
            cAlingnment     =   1
            cBoderStyle     =   2
         End
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 売上実績一覧表 】"
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
         Left            =   11850
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
   End
   Begin VB.Timer Timer2 
      Left            =   705
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   -150
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      Left            =   -15
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8280
      Top             =   7290
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      Caption         =   "開  始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PB_Ok 
      Height          =   585
      Left            =   6915
      Top             =   7290
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      Caption         =   "確　認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin VB.Label PNL_Ok 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6900
      TabIndex        =   16
      Top             =   7965
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8280
      TabIndex        =   13
      Top             =   7965
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5610
      TabIndex        =   12
      Top             =   7965
      Width           =   1485
   End
   Begin VB.Label BG_Logo 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  '透明
      Caption         =   "SIMPR-A"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CXU0120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  売上実績一覧表 発行
'**       フォームID    :  CXU0120
'**       処理概要      :
'**
'**       作  成  日    :  2004/03/23
'**       変  更  日    :  2004/04/22   計算単位による金額再計算をなし。
'**                                     (売上実績作成で再計算済）
'**       変  更  日    :  2004/12/07   頁内最終行と次頁先頭行の品番が異なる時、
'**                                       品番合計が合算されてしまう。
'**       変  更  日    :  2005/03/04 - 重複起動ﾁｪｯｸで起動済の時、『実行時エラー'91'』を表示する対応（Timerﾌﾟﾛﾊﾟﾃｨの Interval=0 に修正）
'**       変  更  日    :  2005/06/09 By CIS 型費の検収日が年月日にて印刷されている不具合を修正
'**       変  更  日    :  2008/06/06 By CIS 工区別に印刷が行えるよう変更
'**                                          カガミ出力
'**       変  更  日    :  2010/01/07 By CIS リスト形式追加
'**       変  更  日    :  2010/02/22 By CIS リスト形式(品番)時SQL修正
'**       変  更  日    :  2011/11/04 By CIS ﾀｲﾑｱｳﾄ延長
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim mPrtDev         As String       ' ﾌﾟﾘﾝﾀ出力先
    Dim mCsvDir         As String
    Dim mCsvKey         As String
    Dim mCsvName        As String
    
    Dim mMsgType        As String
    Dim SL_Select       As String
    
    Dim CisComb1        As CisCombo
    
    Dim mSelColor       As Long
    Dim mUnSelColor     As Long
    Dim mSavePrint      As String
    
    Dim mSaveMsg        As String
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》            +
'+----------------------------------------------------------+
Private Sub Form_Load()
    CisDB.TimeOut = 3000            '2011/11/04
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
    mSaveMsg = lb_msg
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
    
    Set CisComb1 = New CisCombo
    With CisComb1
        Set .Connect = CisDB
        Set .CombObj = H1Comb_Koku
        .Control = "名称マスタ"
        .ComboName = "自社工区"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
    
    With CisFrm
        Set .MSGObject = lb_msg
        Set .MSGTimer = Timer2
        .MSGStopCount = 10
        .MSGTYPE = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

    Dummy.Left = -1000
    Dummy.Enabled = False
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2010/01/07 INSERT START
    H1Op_List(0).Value = True
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2010/01/07 INSERT END
    mSelColor = Op_Out(0).BackColor
    mUnSelColor = Op_Out(1).BackColor
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    
    If gLogoDisp Then
       BG_Logo(0).Caption = gLogo
       CisFrm.BGLDisp BG_Logo, 10
    End If
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》           +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    ' 対象年月初期表示値を求める
    gSL_Select = "Select Max(実績年月) 実績年月 From 売上実績テーブル "
    If UJTRead(gSL_Select, 1) Then
        iH1_YM = UJT.実績年月
    End If

'CSVﾌｧｲﾙ出力先
    mCsvName = "売上実績一覧表"
    mCsvKey = "CXU0120"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", _
                       Key:=mCsvKey, Default:=Mid(CisFun.WinRoot, 1, 3) & mCsvName & ".csv")
''ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始
    
    Call CisComb1.ComboSet
    
    FormAct = Not FormAct
    
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》            +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》    +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function
    
    Select Case Key_Code
        Case vbKeyF7:       If Dummy.Tag <> "" Then Call PB_Look_Click      '【検索】
        Case vbKeyF12
                            If PB_END.Visible Then Call PB_END_Click        '【終了】
                            If PB_Ok.Visible Then Call PB_END_Click         '【終了】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
    '                            Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                Call PB_ENT_Click
    '                            Dummy.Enabled = False
                            End If
        Case Else:          Exit Function
    End Select
    
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                 +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'ｸﾞﾗﾃﾞｰｼｮﾝ表示
    With CisFrm
        .Gradation_Direction = Landscape
        .Gradation_Shading = Lighten
        .Gradation 60, 190, 200, 4, 2, 3
    End With
End Sub
'+----------------------------------------+
'+  Op_Out   出力先                       +
'+----------------------------------------+
Private Sub Op_Out_GotFocus(INDEX As Integer)
    Op_Out(0).Tag = INDEX
    For Each gObj In Op_Out
        If gObj.INDEX = INDEX Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
    Select Case INDEX
        Case 0: lb_PrtDev = mPrtDev
        Case 1: lb_PrtDev = mPrtDev
        Case 2: lb_PrtDev = mCsvDir
        Case Else
            lb_PrtDev = ""
    End Select
'    gp_PrtKB = Index
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2010/01/07 INSERT START
'+----------------------------------------+
'+  H1Op_List   切替                      +
'+----------------------------------------+
Private Sub H1Op_List_GotFocus(INDEX As Integer)
    H1Op_List(0).Tag = INDEX
    For Each gObj In H1Op_List
        If gObj.INDEX = INDEX Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2010/01/07 INSERT END
'+---------------------+
'+    終  了  キ ー    +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp
    
    CisFrm.PrintSet mSavePrint
    
    End
End Sub
'+----------------------------+
'+    確　認  キ ー(F12)      +
'+----------------------------+
Private Sub PB_OK_Click()
    Call PB_END_Click
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)    +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+      検         索         +
'+----------------------------+
Private Sub PB_Look_Click()
    If Dummy.Tag = "iH1_Torcd" Then Call PB_Torcd_Click: Exit Sub
End Sub
'+---------------------+
'+    契約先検索       +
'+---------------------+
Private Sub PB_Torcd_Click()

    RV_TorKb = 2    ' 得意先
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
       iH1_Torcd = RV_TorcdK
       Call iH1_Torcd_LostFocus
    End If

End Sub
'+----------------------------+
'+    iH1_Torcd(GotFocus)     +
'+----------------------------+
Private Sub iH1_Torcd_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+----------------------------+
'+    iH1_Torcd(LostFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_LostFocus()
    Dummy.Tag = ""
    If TorNmGet(iH1_Torcd, 2) Then
        H1lb_Torcd = TRM_RName
        If Trim(H1lb_Torcd) = "" Then
           H1lb_Torcd = TRM_Name
        End If
    Else
        H1lb_Torcd = ""
    End If
End Sub
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理       *
'****************************
Private Function DispChange(Pro As String)
    Dim H1Mode      As Boolean
    Dim B2Mode      As Boolean
    H1Mode = False
    B2Mode = False
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1"
          H1Mode = True
       Case "B2"
          B2Mode = True
    End Select
    H1_Area1.Enabled = H1Mode
    PB_ENT.Visible = H1Mode
    PNL_ENT.Visible = H1Mode
    PB_END.Visible = H1Mode
    PNL_END.Visible = H1Mode
    Timer2.Enabled = H1Mode Or B2Mode
    
    PB_Ok.Visible = B2Mode
    PNL_Ok.Visible = B2Mode
    
    Csv_Bar.Visible = False
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If Not Head1Chk Then GoTo ReturnPress_Ee
        
    Call DispChange("B1")
    
    If Not PrintStart Then GoTo ReturnPress_Ee

    gStr = "【  出  力  終  了  】"
    lb_msg.Caption = gStr
    DoEvents
    
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
    With CisFrm
        .MSGTYPE = FSize_Changes
        .MSGText = gStr
    End With
    ProcHB = "B2"
    Call DispChange(ProcHB)
    Exit Sub
ReturnPress_Ee:
    lb_msg = mSaveMsg
    DoEvents
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理       *
'****************************
Private Function Head1Chk() As Boolean
    Dim wSl_Where   As String
    
    Head1Chk = False
    
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　実績年月を入力して下さい    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If

    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 2) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　契約先　未登録    "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If

    lb_msg.Caption = "【 売上実績データ 抽出中 】"
    DoEvents
'''    '検収　計算方法 = 伝票単位
'''    gSL_Select = "Select UJ.実績年月,UJ.契約先,ISNULL(TR.略称,'') 契約先名," & vbCr
'''    gSL_Select = gSL_Select & "Case When ISNULL(HN.品番,'') = '' Then UJ.品番 Else HN.表示品番 End 品番," & vbCr
'''    gSL_Select = gSL_Select & "納入数 , 単価, 納入番号, 検収年月日, 金額, 仮単価区分" & vbCr
'''    gSL_Select = gSL_Select & " from 売上実績テーブル UJ" & vbCr
'''    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN" & vbCr
'''    gSL_Select = gSL_Select & "   On HN.品番 = UJ.品番" & vbCr
'''    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR" & vbCr
'''    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 0" & vbCr
'''    gSL_Select = gSL_Select & "   And TR.取引先CD = UJ.契約先" & vbCr
'''    gSL_Select = gSL_Select & " Where TR.計算方法 = 0" & vbCr
'''    gSL_Select = gSL_Select & "   And UJ.実績年月 = '" & iH1_YM & "'" & vbCr
'''    gSL_Select = gSL_Select & "   And UJ.入力区分 = ''" & vbCr
'''    If Trim(iH1_Torcd) <> "" Then
'''       gSL_Select = gSL_Select & " And UJ.契約先 = '" & iH1_Torcd & "'" & vbCr
'''    End If
'''    '検収　計算方法 = 月単位
'''    gSL_Select = gSL_Select & " Union ALL " & vbCr
'''    gSL_Select = gSL_Select & "Select UJ.実績年月,UJ.契約先,ISNULL(TR.略称,'') 契約先名," & vbCr
'''    gSL_Select = gSL_Select & "Case When ISNULL(HN.表示品番,'') = '' Then UJ.品番 Else HN.表示品番 End 品番," & vbCr
'''    gSL_Select = gSL_Select & "SUM(納入数) 納入数,単価,'' 納入番号,'' 検収年月日," & vbCr
'''    gSL_Select = gSL_Select & "Case When Isnull(TR.計算単位,0) = 0 Then" & vbCr
'''    gSL_Select = gSL_Select & "          ROUND( SUM(納入数) * 単価,0,1)" & vbCr      '切捨て
'''    gSL_Select = gSL_Select & "     When Isnull(TR.計算単位,0) = 1 Then" & vbCr
'''    gSL_Select = gSL_Select & "          CEILING (Sum(納入数) * 単価)" & vbCr        '切り上げ"
'''    gSL_Select = gSL_Select & "     When Isnull(TR.計算単位,0) = 2 Then" & vbCr
'''    gSL_Select = gSL_Select & "          ROUND( SUM(納入数) * 単価,0)  " & vbCr      ' 四捨五入
'''    gSL_Select = gSL_Select & "     Else" & vbCr
'''    gSL_Select = gSL_Select & "          0 End 金額," & vbCr
'''    gSL_Select = gSL_Select & "仮単価区分" & vbCr
'''    gSL_Select = gSL_Select & " from 売上実績テーブル UJ" & vbCr
'''    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN" & vbCr
'''    gSL_Select = gSL_Select & "   On HN.品番 = UJ.品番" & vbCr
'''    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR" & vbCr
'''    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 0" & vbCr
'''    gSL_Select = gSL_Select & "   And TR.取引先CD = UJ.契約先" & vbCr
'''    gSL_Select = gSL_Select & " Where TR.計算方法 = 1" & vbCr
'''    gSL_Select = gSL_Select & "   And UJ.実績年月 = '" & iH1_YM & "'" & vbCr
'''    gSL_Select = gSL_Select & "   And UJ.入力区分 = ''" & vbCr
'''    If Trim(iH1_Torcd) <> "" Then
'''       gSL_Select = gSL_Select & " And UJ.契約先 = '" & iH1_Torcd & "'" & vbCr
'''    End If
'''    gSL_Select = gSL_Select & " Group by UJ.実績年月,UJ.契約先,ISNULL(TR.略称,''),UJ.品番,HN.表示品番," & vbCr
'''    gSL_Select = gSL_Select & "          単価,仮単価区分,TR.計算単位" & vbCr
'''    '型費・売上入力
'''    gSL_Select = gSL_Select & " Union ALL " & vbCr
'''    gSL_Select = gSL_Select & "Select UJ.実績年月,UJ.契約先,ISNULL(TR.略称,'') 契約先名," & vbCr
'''    gSL_Select = gSL_Select & "Case When ISNULL(HN.品番,'') = '' Then UJ.品番 Else HN.表示品番 End 品番," & vbCr
'''    gSL_Select = gSL_Select & "納入数 , 単価, 納入番号, 検収年月日, 金額, 仮単価区分" & vbCr
'''    gSL_Select = gSL_Select & " from 売上実績テーブル UJ" & vbCr
'''    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN" & vbCr
'''    gSL_Select = gSL_Select & "   On HN.品番 = UJ.品番" & vbCr
'''    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR" & vbCr
'''    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 0" & vbCr
'''    gSL_Select = gSL_Select & "   And TR.取引先CD = UJ.契約先" & vbCr
'''    gSL_Select = gSL_Select & " Where UJ.実績年月 = '" & iH1_YM & "'" & vbCr
'''    gSL_Select = gSL_Select & "   And UJ.入力区分 <> ''" & vbCr
'''    If Trim(iH1_Torcd) <> "" Then
'''       gSL_Select = gSL_Select & " And UJ.契約先 = '" & iH1_Torcd & "'" & vbCr
'''    End If
'2004/04/22変更　再計算は必要なし（売上実績作成で再計算済）
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2010/01/07 INSERT START
    If H1Op_List(0).Value Then
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2010/01/07 INSERT END
        gSL_Select = ""
        gSL_Select = gSL_Select & " Select UJ.実績年月,UJ.契約先,ISNULL(TR.略称,'') 契約先名," & vbCr
        gSL_Select = gSL_Select & "        Case When ISNULL(HN.品番,'') = ''" & vbCr
        gSL_Select = gSL_Select & "              Then UJ.品番 Else HN.表示品番 End 品番," & vbCr
        gSL_Select = gSL_Select & "        納入数 , 単価, 納入番号, 検収年月日, 金額, 仮単価区分" & vbCr
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT START
        gSL_Select = gSL_Select & "      , HT.自社工区, ISNULL(SY1.値名称, '') 自社工区名"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT END
        gSL_Select = gSL_Select & "   from 売上実績テーブル UJ" & vbCr
        gSL_Select = gSL_Select & "        Left Outer Join 品番マスタ HN" & vbCr
        gSL_Select = gSL_Select & "             On HN.品番 = UJ.品番" & vbCr
        gSL_Select = gSL_Select & "        Left Outer Join 取引先マスタ TR" & vbCr
        gSL_Select = gSL_Select & "             ON TR.取引先区分 = 0" & vbCr
        gSL_Select = gSL_Select & "            And TR.取引先CD = UJ.契約先" & vbCr
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT START
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HT" & vbCr
        gSL_Select = gSL_Select & "             ON HT.品番 = UJ.品番" & vbCr
        gSL_Select = gSL_Select & "            AND HT.取引先 = UJ.納入先" & vbCr
        gSL_Select = gSL_Select & "            AND HT.受入 = UJ.受入" & vbCr
        gSL_Select = gSL_Select & "            AND HT.品目 = 0" & vbCr
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY1" & vbCr
        gSL_Select = gSL_Select & "             ON SY1.区分名称 = '自社工区'" & vbCr
        gSL_Select = gSL_Select & "            AND SY1.値 = HT.自社工区" & vbCr
        gSL_Select = gSL_Select & "            AND SY1.区分タイプ = ''" & vbCr
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT END
        gSL_Select = gSL_Select & "  Where UJ.実績年月 = '" & iH1_YM & "'" & vbCr
        If Trim(iH1_Torcd) <> "" Then
           gSL_Select = gSL_Select & " And UJ.契約先 = '" & iH1_Torcd & "'" & vbCr
        End If
        Call CisComb1.CodeGet
        If CisComb1.Code <> "" Then
           gSL_Select = gSL_Select & " And HT.自社工区 = '" & CisComb1.Code & "'" & vbCr
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 UPDATE START
        'gSL_Select = gSL_Select & "Order By UJ.実績年月,UJ.契約先,UJ.品番,UJ.検収年月日,UJ.納入番号,UJ.仮単価区分" & vbCr
        gSL_Select = gSL_Select & "Order By UJ.実績年月, UJ.契約先, HT.自社工区," & vbCr
        gSL_Select = gSL_Select & "         UJ.品番, UJ.検収年月日, UJ.納入番号, UJ.仮単価区分" & vbCr
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 UPDATE END
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2010/01/07 INSERT START
    Else
        gSL_Select = ""
        gSL_Select = gSL_Select & " SELECT UJ.実績年月,UJ.契約先,ISNULL(TR.略称,'') 契約先名," & vbCr
        gSL_Select = gSL_Select & "        CASE WHEN ISNULL(HN.品番,'') = ''" & vbCr
        gSL_Select = gSL_Select & "              THEN UJ.品番 ELSE HN.表示品番 END 品番," & vbCr
        gSL_Select = gSL_Select & "        納入数 , 単価, 納入番号, 検収年月日, 金額, 仮単価区分" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update Start
'        gSL_Select = gSL_Select & "      , HT.自社工区, ISNULL(SY1.値名称, '') 自社工区名" & vbCr
        gSL_Select = gSL_Select & "      , '' 自社工区, '' 自社工区名" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update End
        gSL_Select = gSL_Select & "   FROM (" & vbCr
        gSL_Select = gSL_Select & "          SELECT UT.実績年月,UT.契約先,UT.品番,UT.検収年月日,UT.納入番号" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update Start
        gSL_Select = gSL_Select & "                ,UT.納入数,UT.単価,UT.金額,UT.納入先,UT.仮単価区分" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update End
        gSL_Select = gSL_Select & "            FROM 売上実績テーブル UT" & vbCr
        gSL_Select = gSL_Select & "            LEFT OUTER JOIN (" & vbCr
        gSL_Select = gSL_Select & "                              SELECT 実績年月,契約先,品番,SUBSTRING(検収年月日,1,6) 検収年月" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update Start
'        gSL_Select = gSL_Select & "                                    ,COUNT(*) 件数,納入先,受入,仮単価区分" & vbCr
        gSL_Select = gSL_Select & "                                    ,COUNT(*) 件数,納入先,単価,仮単価区分" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update End
        gSL_Select = gSL_Select & "                                FROM 売上実績テーブル" & vbCr
        gSL_Select = gSL_Select & "                               GROUP BY 実績年月,契約先,品番,SUBSTRING(検収年月日,1,6)" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update Start
'        gSL_Select = gSL_Select & "                                       ,納入先,受入,仮単価区分" & vbCr
        gSL_Select = gSL_Select & "                                       ,納入先,単価,仮単価区分" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update End
        gSL_Select = gSL_Select & "                             ) UL" & vbCr
        gSL_Select = gSL_Select & "              ON UL.実績年月 = UT.実績年月" & vbCr
        gSL_Select = gSL_Select & "             AND UL.契約先 = UT.契約先" & vbCr
        gSL_Select = gSL_Select & "             AND UL.品番 = UT.品番" & vbCr
        gSL_Select = gSL_Select & "             AND UL.検収年月 = SUBSTRING(UT.検収年月日,1,6)" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Insert Start
        gSL_Select = gSL_Select & "             AND UL.納入先 = UT.納入先" & vbCr
        gSL_Select = gSL_Select & "             AND UL.単価 = UT.単価" & vbCr
        gSL_Select = gSL_Select & "             AND UL.仮単価区分 = UT.仮単価区分" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 INsert End
        gSL_Select = gSL_Select & "           WHERE UL.件数 = 1" & vbCr
        gSL_Select = gSL_Select & "          UNION" & vbCr
        gSL_Select = gSL_Select & "          SELECT UL.実績年月,UL.契約先,UL.品番" & vbCr
        gSL_Select = gSL_Select & "                ,CASE WHEN UZ.重複 = 1 THEN '' ELSE UL.検収年月 END 検収年月日" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update Start
'        gSL_Select = gSL_Select & "                ,'' 納入番号,UL.納入数,UL.単価,UL.金額,UL.納入先,UL.受入,UL.仮単価区分" & vbCr
        gSL_Select = gSL_Select & "                ,'' 納入番号,UL.納入数,UL.単価,UL.金額,UL.納入先,UL.仮単価区分" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update End
        gSL_Select = gSL_Select & "            FROM (" & vbCr
        gSL_Select = gSL_Select & "                   SELECT 実績年月,契約先,品番,SUBSTRING(検収年月日,1,6) 検収年月,SUM(納入数) 納入数" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update Start
'        gSL_Select = gSL_Select & "                         ,SUM(単価) 単価,SUM(金額) 金額,COUNT(*) 件数,納入先,受入,仮単価区分" & vbCr
        gSL_Select = gSL_Select & "                         ,単価,SUM(金額) 金額,COUNT(*) 件数,納入先,仮単価区分" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update End
        gSL_Select = gSL_Select & "                     FROM 売上実績テーブル" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update Start
        gSL_Select = gSL_Select & "                    GROUP BY 実績年月,契約先,品番,SUBSTRING(検収年月日,1,6),納入先,単価,仮単価区分" & vbCr
'-----------------------------------------------------------------------------------------------------------------2010/02/22 Update End
        gSL_Select = gSL_Select & "                  ) UL" & vbCr
        gSL_Select = gSL_Select & "            LEFT OUTER JOIN (" & vbCr
        gSL_Select = gSL_Select & "                             SELECT UZ.実績年月,UZ.契約先,UZ.品番,COUNT(品番) 重複" & vbCr
        gSL_Select = gSL_Select & "                               FROM (" & vbCr
        gSL_Select = gSL_Select & "                                     SELECT UZ.実績年月,UZ.契約先,UZ.品番,UZ.検収年月 検収年月日" & vbCr
        gSL_Select = gSL_Select & "                                       FROM (" & vbCr
        gSL_Select = gSL_Select & "                                             SELECT 実績年月,契約先,品番,SUBSTRING(検収年月日,1,6) 検収年月,COUNT(*) 件数" & vbCr
        gSL_Select = gSL_Select & "                                               FROM 売上実績テーブル" & vbCr
        gSL_Select = gSL_Select & "                                              GROUP BY 実績年月,契約先,品番,SUBSTRING(検収年月日,1,6)" & vbCr
        gSL_Select = gSL_Select & "                                             ) UZ" & vbCr
        gSL_Select = gSL_Select & "                                      WHERE UZ.件数 >= 2" & vbCr
        gSL_Select = gSL_Select & "                                     ) UZ" & vbCr
        gSL_Select = gSL_Select & "                              GROUP BY UZ.実績年月,UZ.契約先,UZ.品番" & vbCr
        gSL_Select = gSL_Select & "                             ) UZ" & vbCr
        gSL_Select = gSL_Select & "              ON UZ.実績年月 = UL.実績年月" & vbCr
        gSL_Select = gSL_Select & "             AND UZ.契約先 = UL.契約先" & vbCr
        gSL_Select = gSL_Select & "             AND UZ.品番 = UL.品番" & vbCr
        gSL_Select = gSL_Select & "           WHERE UL.件数 >= 2" & vbCr
        gSL_Select = gSL_Select & "         ) UJ" & vbCr
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HN" & vbCr
        gSL_Select = gSL_Select & "             ON HN.品番 = UJ.品番" & vbCr
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR" & vbCr
        gSL_Select = gSL_Select & "             ON TR.取引先区分 = 0" & vbCr
        gSL_Select = gSL_Select & "            AND TR.取引先CD = UJ.契約先" & vbCr
'-----------------------------------------------------------------------------------2010/02/22 Delete Start
'        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HT" & vbCr
'        gSL_Select = gSL_Select & "             ON HT.品番 = UJ.品番" & vbCr
'        gSL_Select = gSL_Select & "            AND HT.取引先 = UJ.納入先" & vbCr
'        gSL_Select = gSL_Select & "            AND HT.受入 = UJ.受入" & vbCr
'        gSL_Select = gSL_Select & "            AND HT.品目 = 0" & vbCr
'        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY1" & vbCr
'        gSL_Select = gSL_Select & "             ON SY1.区分名称 = '自社工区'" & vbCr
'        gSL_Select = gSL_Select & "            AND SY1.値 = HT.自社工区" & vbCr
'        gSL_Select = gSL_Select & "            AND SY1.区分タイプ = ''" & vbCr
'-----------------------------------------------------------------------------------2010/02/22 Delete End
        gSL_Select = gSL_Select & "  WHERE UJ.実績年月 = '" & iH1_YM & "'" & vbCr
        If Trim(iH1_Torcd) <> "" Then
           gSL_Select = gSL_Select & " AND UJ.契約先 = '" & iH1_Torcd & "'" & vbCr
        End If
'-----------------------------------------------------------------------------------2010/02/22 Delete Start
'        Call CisComb1.CodeGet
'        If CisComb1.Code <> "" Then
'           gSL_Select = gSL_Select & " AND HT.自社工区 = '" & CisComb1.Code & "'" & vbCr
'        End If
'-----------------------------------------------------------------------------------2010/02/22 Delete End
'-----------------------------------------------------------------------------------2010/02/22 Update Start
'        gSL_Select = gSL_Select & "ORDER BY UJ.実績年月, UJ.契約先, HT.自社工区," & vbCr
        gSL_Select = gSL_Select & "ORDER BY UJ.実績年月, UJ.契約先," & vbCr
'-----------------------------------------------------------------------------------2010/02/22 Update End
        gSL_Select = gSL_Select & "         UJ.品番, UJ.検収年月日, UJ.納入番号, UJ.仮単価区分" & vbCr
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2010/01/07 INSERT END
    
    If Not UJTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　該当データがありません       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    
    Head1Chk = True
End Function

'*******************************
'*      タイトルバー編集       *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      メッセージ表示         *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
End Sub
'+-------------------------------+
'+     プリンタ・CSV出力先 変更  +
'+-------------------------------+
Private Sub PB_Print_Click()
    If Op_Out(2).Value Then
        ' CSV出力先変更
        On Error Resume Next
        CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly
        CommonDialog1.FileName = mCsvDir
        CommonDialog1.Filter = "ＣＳＶ ファイル (*.csv)|*.csv"
        ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
        CommonDialog1.CancelError = True
        CommonDialog1.ShowOpen
        If Err = 32755 Then
            '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
            '必要なｿｰｽｺｰﾄﾞを入力してください。
        End If
        If Trim(CommonDialog1.FileName) <> "" Then
            mCsvDir = CommonDialog1.FileName
            lb_PrtDev.Caption = mCsvDir
        End If
        On Error GoTo 0
    Else
    'If Op_Out(0).Value Or Op_Out(1).Value Then
        ' プリンタ種類変更
         On Error Resume Next
         CommonDialog1.Flags = 64 + cdlPDHidePrintToFile
         CommonDialog1.Max = 0
         CommonDialog1.Min = 0
         CommonDialog1.Copies = 1
         CommonDialog1.ToPage = 1
         CommonDialog1.FromPage = 1
        ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
         CommonDialog1.CancelError = True
         CommonDialog1.ShowPrinter
         If Err = 32755 Then
        '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
        '必要なｿｰｽｺｰﾄﾞを入力してください。
         Else
             mPrtDev = Printer.DeviceName
             lb_PrtDev.Caption = mPrtDev
         End If
         On Error GoTo 0
    End If
End Sub
'+---------------------------------------------+
'+      印刷処理                               +
'+---------------------------------------------+
Private Function PrintStart() As Boolean
    PrintStart = False
    If Op_Out(0).Value Or Op_Out(1).Value Then
        lb_msg.Caption = "【 売上実績一覧表 印刷中 】"
        DoEvents
        With CisPrt
             Set .iPrtForm = CXU0120L
             .dpMSG = "売上実績一覧表　印刷中"
             If Op_Out(0).Value Then
                .iPreview = pv_Print
             Else
                .iPreview = pv_Preview
             End If
             .PrintStart
        End With
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT START
'        RV_YM = RTrim(iH1_YM)
'        RV_Torcd = RTrim(iH1_Torcd)
'        Call CisComb1.CodeGet
'        RV_KOKBN = RTrim(CisComb1.Code)
'        lb_msg.Caption = "【 売上実績表(カガミ) 印刷中 】"
'        DoEvents
'        With CisPrt
'             Set .iPrtForm = CXU0125L
'             .dpMSG = "売上実績表(カガミ)　印刷中"
'             If Op_Out(0).Value Then
'                .iPreview = pv_Print
'             Else
'                .iPreview = pv_Preview
'             End If
'             .PrintStart
'        End With
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT END
    Else
        If Not CsvPut Then Exit Function
    End If
    PrintStart = True
End Function
'****************************
'*       CSV出力処理        *
'****************************
Private Function CsvPut() As Boolean

    Dim OutStr      As String
    Dim OutFNo      As Long
    Dim OutCnt      As Long
    
    CsvPut = False
    
    ' 出力ﾌｧｲﾙ存在確認
    If Dir(mCsvDir) <> "" Then
        DoEvents
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　 " & mCsvDir & "　ﾌｧｲﾙは既に存在します。                 "
            .MB_MSG(4) = "   　 " & "上書きしますか？           "
            .MB_Title = "上書き確認"
            .MB_Button = Yes_No
            If Not .MBOX Then Exit Function
        End With
    End If
    
    lb_msg.Caption = "【 売上実績情報 出力中 】"
    Csv_Bar.Visible = True
    
    DoEvents
        
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo

    Call CisFrm.MousePT(11)
    
    With CisFrm
        Set .MeterPic = Csv_Bar
        .MeterMax = CisDB.RecordCount
    End With
        
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 UPDATE START
    'OutStr = "№,実績年月,契約先,契約先名,納入番号,品番,納入数,単価,金額,検収日,仮単価 "
    OutStr = "№,実績年月,契約先,契約先名,自社工区,自社工区名,納入番号,品番,納入数,単価,金額,検収日,仮単価 "
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 UPDATE END
    Print #OutFNo, OutStr
    OutCnt = 0
    Do Until Not UJT_RDSTS
        With UJT
            OutCnt = OutCnt + 1
            OutStr = OutCnt & ","
            OutStr = OutStr & Mid(.実績年月, 1, 4) & "/" & Mid(.実績年月, 5, 2) & ","
            OutStr = OutStr & RTrim(.契約先) & ","
            OutStr = OutStr & RTrim(.契約先名) & ","
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT START
            OutStr = OutStr & RTrim(.自社工区) & ","
            OutStr = OutStr & RTrim(.自社工区名) & ","
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT END
            OutStr = OutStr & RTrim(.納入番号) & ","
            OutStr = OutStr & RTrim(.品番) & ","
            OutStr = OutStr & .納入数 & ","
            OutStr = OutStr & .単価 & ","
            OutStr = OutStr & .金額 & ","
            If Trim(.検収年月日) <> "" Then
'< 2005/06/09 Delete START >-----------------------------------------------------------------------
'                With CisFun
'                    .DateE_BefVal = UJT.検収年月日
'                    .DateEdit
'                    OutStr = OutStr & .DateE_AfVal
'                End With
'< 2005/06/09 Delete End >-----------------------------------------------------------------------
'< 2005/06/09 ADD START >-----------------------------------------------------------------------
                With CisFun
                    .DateE_BefVal = UJT.検収年月日
                    If Len(UJT.検収年月日) > 6 Then
                       .DateEdit
                       OutStr = OutStr & "'" & .DateE_AfVal
                    Else
                       OutStr = OutStr & "'" & Mid(UJT.検収年月日, 1, 4) & "/" & Mid(UJT.検収年月日, 5, 2)
                    End If
                End With
'< 2005/06/09 ADD End >-----------------------------------------------------------------------
            Else
                OutStr = OutStr & ","
            End If
            If Trim(.仮単価区分) <> "" Then
                OutStr = OutStr & "*" & ""
            End If

            Print #OutFNo, OutStr
    
            CisFrm.MeterUpdate OutCnt

            Call UJTReadNext
        End With
    Loop

CsvPut_Ed:
    Call UJTClose

    Close #OutFNo
    SaveSetting appname:=gRegKey, section:="CSV", _
                    Key:=mCsvKey, setting:=RTrim(lb_PrtDev)
    
    Call CisFrm.MousePT(1)
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "   出力処理が終了しました。    "
        .MB_MSG(4) = "   　 " & "出力先    = " & mCsvDir
        .MB_MSG(5) = "   　 " & "出力件数 = " & StrConv(Format(CisDB.RecordCount, "##,##0"), vbWide) & " 件 "
        .MB_Title = "出力確認"
        .MB_Button = OK
        .MBOX
    End With
    CsvPut = True
    
End Function

