VERSION 5.00
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CZK0150 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料棚卸リスト発行"
   ClientHeight    =   11190
   ClientLeft      =   2040
   ClientTop       =   3405
   ClientWidth     =   15390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15390
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4260
      ScaleHeight     =   285
      ScaleWidth      =   6600
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2730
      Visible         =   0   'False
      Width           =   6660
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5625
      Top             =   7845
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
      Height          =   4275
      Left            =   4260
      ScaleHeight     =   4387.046
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6585
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3135
      Width           =   6645
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   4140
         Index           =   1
         Left            =   60
         Top             =   45
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   60
            TabIndex        =   15
            Top             =   3000
            Width           =   6345
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1065
               TabIndex        =   5
               Top             =   270
               Value           =   -1  'True
               Width           =   1260
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   2505
               TabIndex        =   6
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   4380
               TabIndex        =   7
               Top             =   240
               Width           =   975
            End
         End
         Begin CisText_V60.CisText iH1_Basyo 
            Height          =   360
            Left            =   1800
            TabIndex        =   2
            Top             =   1545
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XX"
            MaxLength       =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_Basyonm 
            Height          =   360
            Left            =   2565
            Top             =   1545
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
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
               TabIndex        =   9
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
                  Caption         =   "材料棚卸リストの発行を行います"
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
            Top             =   3630
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
            Top             =   3660
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
            Height          =   360
            Index           =   1
            Left            =   405
            Top             =   900
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "棚卸年月"
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   300
            Index           =   0
            Left            =   540
            Top             =   1560
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "在庫場所"
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
         Begin Cis3D_v60.CIS3D PB_Basyo 
            Height          =   360
            Left            =   2220
            Top             =   1545
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
            Height          =   360
            Left            =   3720
            Top             =   2250
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "～"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cAlingnment     =   1
            cBoderStyle     =   2
            cPositionX      =   20
         End
         Begin CisYMD_V60.CisYMD iH1_SYMD 
            Height          =   375
            Left            =   1785
            TabIndex        =   3
            Top             =   2235
            Width           =   1875
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
            Object.Height          =   375
            Object.Width           =   1875
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   300
            Index           =   1
            Left            =   540
            Top             =   2280
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "最終出庫日"
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
            Left            =   1800
            TabIndex        =   1
            Top             =   900
            Width           =   1470
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
         Begin CisYMD_V60.CisYMD iH1_EYMD 
            Height          =   375
            Left            =   4050
            TabIndex        =   4
            Top             =   2235
            Width           =   1875
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
            Object.Height          =   375
            Object.Width           =   1875
            cSize           =   -1  'True
            cChkResult      =   0   'False
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
      Caption         =   "【 材料棚卸リスト発行 】"
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
      Top             =   7845
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
      Top             =   7845
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4260
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label PNL_Ok 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6900
      TabIndex        =   14
      Top             =   8520
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8280
      TabIndex        =   12
      Top             =   8520
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5610
      TabIndex        =   11
      Top             =   8520
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
      TabIndex        =   10
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CZK0150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  材料棚卸リスト 発行
'**       フォームID    :  CZK0150
'**       処理概要      :
'**
'**       作  成  日    :  2010/01/15
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
'
    Dim mSaveMsg        As String       ' ﾒｯｾｰｼﾞ
    Dim mSavePrint      As String
    Dim mPrtDev         As String       ' ﾌﾟﾘﾝﾀ出力先
    Dim mCsvKey         As String
    Dim mCsvDir         As String
    Dim mCsvName        As String

'   背景色
    Dim mSelColor       As Long
    Dim mUnSelColor     As Long
    Dim mUnKubunColor   As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub form_load()
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

    Dummy.Left = -1000
    Dummy.Enabled = False

    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始

    With CisFrm
        Set .MSGObject = lb_msg
        Set .MSGTimer = Timer2
        .MSGStopCount = 10
        .MSGTYPE = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

'   背景色の退避
    mSelColor = Op_Out(0).BackColor
    mUnSelColor = Op_Out(1).BackColor

    If gLogoDisp Then
       BG_Logo(0).Caption = gLogo
       CisFrm.BGLDisp BG_Logo, 10
    End If

    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    'Call CisComb1.ComboSet
    
'   CSVﾌｧｲﾙ出力先
    mCsvName = Mid(CisFun.WinRoot, 1, 3) & "材料棚卸リスト.csv"
    mCsvKey = "CZK0150"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", Key:=mCsvKey, Default:=mCsvName)
'   ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev

    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
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
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'ｸﾞﾗﾃﾞｰｼｮﾝ表示
    With CisFrm
        .Gradation_Direction = Landscape
        .Gradation_Shading = Lighten
        .Gradation 100, 130, 200, 4, 2, 3
    End With
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp

    CisFrm.PrintSet mSavePrint

    End
End Sub
'+----------------------------+
'+    確　認  キ ー(F12)       +
'+----------------------------+
Private Sub PB_OK_Click()
    Call PB_END_Click
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
'+      検         索         +
'+----------------------------+
Private Sub PB_Look_Click()
    If Dummy.Tag = "iH1_Basyo" Then Call PB_Basyo_Click: Exit Sub
End Sub
'+---------------------+
'+    在庫場所検索     +
'+---------------------+
Private Sub PB_Basyo_Click()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "管理場所"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iH1_Basyo = RV_Code
        Call iH1_Basyo_LostFocus
    End If
    
End Sub
'+----------------------------------------+
'+  Op_Out   出力先
'+----------------------------------------+
Private Sub Op_Out_GotFocus(Index As Integer)
    Op_Out(0).Tag = Index
    For Each gObj In Op_Out
        If gObj.Index = Index Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
    Select Case Index
        Case 0: lb_PrtDev = mPrtDev
        Case 1: lb_PrtDev = mPrtDev
        Case 2: lb_PrtDev = mCsvDir
        Case Else
            lb_PrtDev = ""
    End Select
End Sub
'+----------------------------+
'+    iH1_Basyo(GotFocus)    +
'+----------------------------+
Private Sub iH1_Basyo_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+----------------------------+
'+    iH1_Basyo(LostFocus)    +
'+----------------------------+
Private Sub iH1_Basyo_LostFocus()
    Dummy.Tag = ""
    
    H1lb_Basyonm = ""
    If MeisyoGet("管理場所", RTrim(iH1_Basyo)) Then
        H1lb_Basyonm = RTrim(SYM_Meisyo)
    End If
    
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      メッセージ表示        *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
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
       Case "H1":   H1Mode = True
       Case "B2":   B2Mode = True
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

    If Not Head1Chk Then GoTo ReturnPress_Ed

    Call DispChange("B1")
    
    If Not Create_List Then GoTo ReturnPress_Ee

    If Not PrintStart Then GoTo ReturnPress_Ee

    lb_msg.Caption = "【  出  力  終  了  】"
    DoEvents

    Call CisFrm.MousePT(1)
    With CisFrm
        .MSGTYPE = FSize_Changes
        .MSGText = "【  出  力  終  了  】"
    End With
    ProcHB = "B2"
    Call DispChange(ProcHB)
    Exit Sub
ReturnPress_Ee:
    lb_msg = mSaveMsg
    DoEvents
    ProcHB = "H1"
    Call DispChange(ProcHB)
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

'   開始日
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    棚卸日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    
    If Mid(iH1_YM, 5, 2) < 1 Or Mid(iH1_YM, 5, 2) > 12 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    棚卸日を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If

'   在庫場所
    If Trim(iH1_Basyo) <> "" Then
        If Not MeisyoGet("管理場所", RTrim(iH1_Basyo)) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    在庫場所 未登録         "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
            iH1_Basyo.SetFocus
            Exit Function
        End If
    End If
    
'   最終出庫日
    If RTrim(iH1_SYMD) <> "" Then
        If iH1_SYMD.cYear = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終出庫日 - 年 を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_SYMD.SetFocus
           Exit Function
        End If
        
        If iH1_SYMD.cMonth = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終出庫日 - 月 を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_SYMD.SetFocus
           Exit Function
        End If
        
        If iH1_SYMD.cDay = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終出庫日 - 日 を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_SYMD.SetFocus
           Exit Function
        End If
        
        If iH1_SYMD.cChkResult = False Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    入力された最終出庫日は日付として正しくありません。       "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_SYMD.SetFocus
           Exit Function
        End If
    End If
    If RTrim(iH1_EYMD) <> "" Then
        If iH1_EYMD.cYear = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終出庫日 - 年 を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_EYMD.SetFocus
           Exit Function
        End If
        
        If iH1_EYMD.cMonth = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終出庫日 - 月 を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_EYMD.SetFocus
           Exit Function
        End If
        
        If iH1_EYMD.cDay = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終出庫日 - 日 を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_EYMD.SetFocus
           Exit Function
        End If
        
        If iH1_EYMD.cChkResult = False Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    入力された最終出庫日は日付として正しくありません。       "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_EYMD.SetFocus
           Exit Function
        End If
    End If
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "   　材料棚卸リスト発行を行います。            "
        .MB_MSG(3) = "   　よろしいですか？            "
        .MB_Title = "材料棚卸リスト発行"
        .MB_Button = Yes_No
        If Not .MBOX Then Exit Function
    End With
    
    'DoEvents
    
    Head1Chk = True
End Function

'+-------------------------------+
'+     プリンタ・CSV出力先 変更    +
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
        lb_msg.Caption = "【 材料棚卸リスト 印刷中 】"
        DoEvents
        With CisPrtCZK0150
            .cDataMaxCnt = CisDB.RecordCount
            Set .iPrtForm = CZK0150L
            .dpMSG = "材料棚卸リスト　印刷中"
            If Op_Out(0).Value Then
                .iPreview = pv_Print
            Else
                .iPreview = pv_Preview
            End If
            .PrintStart
        End With
    Else
        If Not CsvPut Then Exit Function
    End If
    PrintStart = True
End Function
'****************************
'*       CSV出力処理         *
'****************************
Private Function CsvPut() As Boolean
    Dim OutFNo      As Long
    Dim OutStr      As String
    Dim OutCnt      As Long

    CsvPut = False

    On Error GoTo CsvPut_Err

'   出力ﾌｧｲﾙ存在確認
    If Dir(mCsvDir) <> "" Then
        DoEvents
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "   　" & mCsvDir & "　ﾌｧｲﾙは既に存在します。            "
            .MB_MSG(3) = "   　" & "上書きしますか？            "
            .MB_Title = "上書き確認"
            .MB_Button = Yes_No
            If Not .MBOX Then Exit Function
        End With
    End If

    lb_msg.Caption = "【 材料棚卸リストデータ 出力中 】"
    Csv_Bar.Visible = True

    DoEvents

    Call CisFrm.MousePT(11)

    With CisFrm
        Set .MeterPic = Csv_Bar
        .MeterMax = CisDB.RecordCount
    End With

'   >> 出力ﾃｷｽﾄ Open
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo
'   >> Head部 編集＆出力
    Call CsvEdit_Head1(OutStr)
    Print #OutFNo, OutStr
    Call CsvEdit_Head2(OutStr)
    Print #OutFNo, OutStr
'   >> ｸﾘｱ
    OutCnt = 0
    Do Until Not ZTM_RDSTS
    '   >> ｶｳﾝﾄ
        OutCnt = OutCnt + 1
    '   >> Body部 編集＆出力
        Call CsvEdit_Body1(OutStr, OutCnt)
        Print #OutFNo, OutStr

        CisFrm.MeterUpdate OutCnt

    '   >> 次のﾚｺｰﾄﾞを読み込む
        Call ZTMReadNext
    Loop
    Call ZTMClose
'   >> 出力ﾃｷｽﾄ Close
    Close #OutFNo

    SaveSetting appname:=gRegKey, section:="CSV", Key:=mCsvKey, setting:=RTrim(mCsvDir)

    Call CisFrm.MousePT(1)

    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "    【  CSV出力 正常終了 】            "
        .MB_MSG(3) = "    出 力 先  = " & RTrim(mCsvDir) & "　　　　　"
        .MB_MSG(4) = "    出力件数 = " & Format(OutCnt, "#,##0") & " 件"
        .MB_Title = "【確認】"
        .MB_Button = OK
        .MBOX
    End With

    CsvPut = True
    Exit Function
CsvPut_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------+
'+    Head部　編集処理
'+----------------------------+
Private Sub CsvEdit_Head1(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & Mid(RTrim(iH1_YM), 1, 4) & "年" & Mid(RTrim(iH1_YM), 5, 2) & "月" & "棚卸リスト"
End Sub
Private Sub CsvEdit_Head2(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & "在庫場所"
    CsvStr = CsvStr & ",手配先"
    CsvStr = CsvStr & ",棚卸年月日"
    CsvStr = CsvStr & ",材質 / 寸法"
    CsvStr = CsvStr & ",材料名"
    CsvStr = CsvStr & ",仕様"
    CsvStr = CsvStr & ",棚卸数"
    CsvStr = CsvStr & ","
    CsvStr = CsvStr & ",単価"
    CsvStr = CsvStr & ",金額"
    CsvStr = CsvStr & ",最終出庫日"
End Sub
'+----------------------------+
'+    Body部　編集処理
'+----------------------------+
Private Sub CsvEdit_Body1(CsvStr As String, CsvCnt As Long)
    
    Dim WTanka As Currency
    
    With ZTM
        CsvStr = ""
        CsvStr = CsvStr & RTrim(.在庫場所)
        CsvStr = CsvStr & "," & RTrim(.取引先名)
        CsvStr = CsvStr & "," & RTrim(.実棚日)
        CsvStr = CsvStr & "," & RTrim(.材寸)
        CsvStr = CsvStr & "," & RTrim(.材料名)
        CsvStr = CsvStr & "," & RTrim(.仕様)
        CsvStr = CsvStr & "," & Format(.在庫数, "########0")
        CsvStr = CsvStr & "," & RTrim(.手配単位名)
        
        ' 単価算出及び金額計算
        With CisDB
            .SQL = "材料購入単価決定"
            .StoadoCount = 9
            .ParaNo = 0:    .ParaIO = Return用
            .ParaNo = 1:    .ParaIO = Input用
            .ParaNo = 2:    .ParaIO = Input用
            .ParaNo = 3:    .ParaIO = Input用
            .ParaNo = 4:    .ParaIO = Input用
            .ParaNo = 5:    .ParaIO = OutPut用
            .ParaNo = 6:    .ParaIO = OutPut用
            .ParaNo = 7:    .ParaIO = OutPut用
            .ParaNo = 8:    .ParaIO = OutPut用
            .ParaNo = 9:    .ParaIO = OutPut用

            .ParaNo = 1:    .ParaValue = ZTM.材料管理番号               ' 材料管理番号
            .ParaNo = 2:    .ParaValue = RTrim(ZTM.取引先)              ' 取引先
            .ParaNo = 3:    .ParaValue = ""                             ' 契約先
            If Len(RTrim(ZTM.実棚日)) = 6 Then
                .ParaNo = 4:    .ParaValue = "01"
            Else
                .ParaNo = 4:    .ParaValue = RTrim(ZTM.実棚日)          ' 単価決定日付
            End If
            
            .DBStored

           .ParaNo = 0
           If .ParaValue <> 0 Then
               CisFun.MB_Lines = 5
               CisFun.MB_MSG(2) = "    材料購入単価決定　異常！      "
               CisFun.MB_MSG(4) = "    Error No = " & .ParaValue
               CisFun.MB_Button = Error
               CisFun.MBOX
               End
           End If

            .ParaNo = 5
            WTanka = 0
            WTanka = RTrim(.ParaValue)

        End With
        
        CsvStr = CsvStr & "," & Format(WTanka, "########0.00")
        CsvStr = CsvStr & "," & Format((WTanka * .在庫数), "########0")
        CsvStr = CsvStr & "," & RTrim(.最終出庫日)
    End With
End Sub
'/////////////////////////////////////////////////////
'+---------------------------------------------------+
'+               棚卸リストデータ作成
'+---------------------------------------------------+
'/////////////////////////////////////////////////////
Private Function Create_List() As Boolean
    Create_List = False
    
    lb_msg.Caption = "【 棚卸リストデータ 抽出中 】"
    DoEvents
    
'   < 印刷データ読み込み >
    gSL_Select = ""
    gSL_Select = gSL_Select & "Select DT.在庫場所,DT.取引先名,DT.実棚日,DT.材寸,DT.材料名,DT.仕様," & Chr(13)
    gSL_Select = gSL_Select & "       DT.在庫数,DT.手配単位名,DT.最終出庫日,DT.材料管理番号,DT.取引先"
    gSL_Select = gSL_Select & "  From (" & Chr(13)
    gSL_Select = gSL_Select & "         Select ISNULL(MY.値名称,'') 在庫場所,IsNull(TR.略称,'') 取引先名," & Chr(13)
    gSL_Select = gSL_Select & "                Case When IsNull(ZZM.実棚日,'') <> '' then ZZM.実棚日" & Chr(13)
    gSL_Select = gSL_Select & "                Else ZZM.年月度 End 実棚日," & Chr(13)
    gSL_Select = gSL_Select & "                ZR.材質 + ' ' + ZR.表示寸法 材寸,ISNULL(ZR.材料名,'') 材料名," & Chr(13)
    gSL_Select = gSL_Select & "                ISNULL(MT.値名称,'') 仕様,ZZM.在庫数,ISNULL(MS.値名称,'') 手配単位名," & Chr(13)
    'gSL_Select = gSL_Select & "                ISNULL(JK.実績日,'') 実績日,ISNULL(ZS.作業日,'') 作業日," & Chr(13)
    gSL_Select = gSL_Select & "                Case When ISNULL(JK.実績日,'') > ISNULL(ZS.作業日,'') then ISNULL(JK.実績日,'')" & Chr(13)
    gSL_Select = gSL_Select & "                     When ISNULL(JK.実績日,'') < ISNULL(ZS.作業日,'') then ISNULL(ZS.作業日,'')" & Chr(13)
    gSL_Select = gSL_Select & "                     Else ISNULL(JK.実績日,'') End 最終出庫日,ZZM.材料管理番号,ZT.取引先" & Chr(13)
    gSL_Select = gSL_Select & "           From 材料取引先マスタ ZT" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer Join 材料在庫マスタ ZZM" & Chr(13)
    gSL_Select = gSL_Select & "             On ZZM.年月度       = '" & iH1_YM & "'" & Chr(13)
    gSL_Select = gSL_Select & "            And ZZM.材料管理番号 = ZT.材料管理番号" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer Join 取引先マスタ TR" & Chr(13)
    gSL_Select = gSL_Select & "             On TR.取引先CD = ZT.取引先" & Chr(13)
    gSL_Select = gSL_Select & "            And TR.取引先区分 = '1'" & Chr(13)
'    gSL_Select = gSL_Select & "            And TR.契約先区分 = '0'" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer Join 材料マスタ ZR" & Chr(13)
    gSL_Select = gSL_Select & "             On ZR.材料管理番号 = ZT.材料管理番号" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer Join 名称マスタ MY" & Chr(13)
    gSL_Select = gSL_Select & "             On MY.区分名称 = '管理場所'" & Chr(13)
    gSL_Select = gSL_Select & "            And MY.値       = SubString(ZT.置場,1,2)" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer Join 名称マスタ MT" & Chr(13)
    gSL_Select = gSL_Select & "             On MY.区分名称 = '管理場所ユーザー'" & Chr(13)
    gSL_Select = gSL_Select & "            And MY.値       = SubString(ZT.置場,3,2)" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer Join 名称マスタ MS" & Chr(13)
    gSL_Select = gSL_Select & "             On MS.区分名称 = '単位区分'" & Chr(13)
    gSL_Select = gSL_Select & "            And MS.値       = ZR.手配単位" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer Join (Select 実績区分,入力種類,材質,板厚,幅,長さ,Max(実績日) 実績日" & Chr(13)
    gSL_Select = gSL_Select & "                              from 材料実績管理テーブル" & Chr(13)
    gSL_Select = gSL_Select & "                              Group By 実績区分,入力種類,材質,板厚,幅,長さ) JK" & Chr(13)
    gSL_Select = gSL_Select & "             On JK.材質 = ZT.材質" & Chr(13)
    gSL_Select = gSL_Select & "            And JK.板厚 = ZT.板厚" & Chr(13)
    gSL_Select = gSL_Select & "            And JK.幅   = ZT.幅" & Chr(13)
    gSL_Select = gSL_Select & "            And JK.長さ = ZT.長さ" & Chr(13)
    gSL_Select = gSL_Select & "            And JK.実績区分 <> 1" & Chr(13)
    gSL_Select = gSL_Select & "            And JK.入力種類 <> 'A'" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer join (select 材料管理番号,MAX(作業日) 作業日" & Chr(13)
    gSL_Select = gSL_Select & "                              from 材料使用量テーブル" & Chr(13)
    gSL_Select = gSL_Select & "                             Group By 材料管理番号)  ZS" & Chr(13)
    gSL_Select = gSL_Select & "             On ZS.材料管理番号 = ZT.材料管理番号"
    '棚卸日
    gSL_Select = gSL_Select & "          WHere ZZM.年月度 = '" & iH1_YM & "'" & Chr(13)
    '在庫場所
    If RTrim(iH1_Basyo) <> "" Then
        gSL_Select = gSL_Select & "             And SY.値       = '" & RTrim(iH1_Basyo) & "'" & Chr(13)
    End If
    '最終出庫日
    If RTrim(iH1_SYMD) <> "" And RTrim(iH1_EYMD) <> "" Then
        gSL_Select = gSL_Select & "             And ((Case When ISNULL(JK.実績日,'') > ISNULL(ZS.作業日,'') then ISNULL(JK.実績日,'')" & Chr(13)
        gSL_Select = gSL_Select & "                        When ISNULL(JK.実績日,'') < ISNULL(ZS.作業日,'') then ISNULL(ZS.作業日,'')" & Chr(13)
        gSL_Select = gSL_Select & "                        Else ISNULL(JK.実績日,'') End ) < '" & iH1_SYMD & "'"
        gSL_Select = gSL_Select & "              Or  (Case When ISNULL(JK.実績日,'') > ISNULL(ZS.作業日,'') then ISNULL(JK.実績日,'')" & Chr(13)
        gSL_Select = gSL_Select & "                        When ISNULL(JK.実績日,'') < ISNULL(ZS.作業日,'') then ISNULL(ZS.作業日,'')" & Chr(13)
        gSL_Select = gSL_Select & "                        Else ISNULL(JK.実績日,'') End ) > '" & iH1_EYMD & "')"
    Else
        If RTrim(iH1_SYMD) <> "" Then
            gSL_Select = gSL_Select & "             And (Case When ISNULL(JK.実績日,'') > ISNULL(ZS.作業日,'') then ISNULL(JK.実績日,'')" & Chr(13)
            gSL_Select = gSL_Select & "                       When ISNULL(JK.実績日,'') < ISNULL(ZS.作業日,'') then ISNULL(ZS.作業日,'')" & Chr(13)
            gSL_Select = gSL_Select & "                       Else ISNULL(JK.実績日,'') End ) < '" & iH1_SYMD & "'"
        End If
        If RTrim(iH1_EYMD) <> "" Then
            gSL_Select = gSL_Select & "             And (Case When ISNULL(JK.実績日,'') > ISNULL(ZS.作業日,'') then ISNULL(JK.実績日,'')" & Chr(13)
            gSL_Select = gSL_Select & "                       When ISNULL(JK.実績日,'') < ISNULL(ZS.作業日,'') then ISNULL(ZS.作業日,'')" & Chr(13)
            gSL_Select = gSL_Select & "                       Else ISNULL(JK.実績日,'') End ) > '" & iH1_EYMD & "'"
        End If
    End If
    gSL_Select = gSL_Select & "       ) DT" & Chr(13)
    gSL_Select = gSL_Select & " Order by DT.在庫場所,DT.取引先名,DT.実棚日"
    
    If Not ZTMRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　該当データはありません。            "
            .MB_Title = "材料棚卸リスト発行"
            .MB_Button = Error
            If Not .MBOX Then Exit Function
        End With
       Call ZZMClose
       Exit Function
    End If
    
    Create_List = True
End Function



