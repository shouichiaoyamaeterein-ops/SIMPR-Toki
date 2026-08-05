VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CNJ1130 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "工事手配書発行"
   ClientHeight    =   9420
   ClientLeft      =   3090
   ClientTop       =   2805
   ClientWidth     =   14520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9420
   ScaleWidth      =   14520
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4260
      ScaleHeight     =   285
      ScaleWidth      =   6600
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   6660
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
      Top             =   6855
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
      Height          =   3690
      Left            =   4260
      ScaleHeight     =   3778.168
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2565
      Width           =   6660
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   3510
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   6191
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
            Caption         =   "出力先"
            Height          =   600
            Left            =   75
            TabIndex        =   12
            Top             =   2430
            Width           =   6345
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   4380
               TabIndex        =   5
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   2505
               TabIndex        =   4
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1065
               TabIndex        =   3
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
               TabIndex        =   8
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
                  Caption         =   "工事手配書の発行を行います"
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
            Top             =   3060
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
            Top             =   3090
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
            Left            =   195
            Top             =   945
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   1575
            TabIndex        =   2
            Top             =   1800
            Width           =   990
            _ExtentX        =   1746
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   2895
            Top             =   1800
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   300
            Index           =   1
            Left            =   300
            Top             =   1815
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D PB_Torcd 
            Height          =   360
            Left            =   2550
            Top             =   1800
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
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   360
            Left            =   1590
            TabIndex        =   0
            Top             =   960
            Width           =   1560
            _ExtentX        =   2646
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
            cContent        =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   360
            Index           =   0
            Left            =   3240
            Top             =   945
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "計画年月"
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
         Begin CisYM_V60.CisYM iH1_NYM 
            Height          =   360
            Left            =   4635
            TabIndex        =   1
            Top             =   960
            Width           =   1560
            _ExtentX        =   2646
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
            cContent        =   1
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
      Caption         =   "【 工事手配書発行 】"
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
      TabIndex        =   7
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8280
      Top             =   6855
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
      Top             =   6855
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
      TabIndex        =   14
      Top             =   7530
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8280
      TabIndex        =   11
      Top             =   7530
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5610
      TabIndex        =   10
      Top             =   7530
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
      TabIndex        =   9
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CNJ1130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  工事手配書発行
'**       フォームID    :  CNJ1130
'**       処理概要      :
'**
'**       作  成  日    :  2010/01/17
'**       変  更  日    :  2010/11/11  品番取引先マスタから出力,計画年月入力
'**       変  更  日    :  2011/11/28  品番=親品番の場合､出力されるように変更
'**       変  更  日    :  2012/11/26  受注数０の場合も出力
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
    
    Dim wZYM            As String
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
    gConnectCount = 3
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
    
'年月初期値･･･展開処理最新年月
    gSL_Select = "SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE 区分名称 = '処理内容'"
    gSL_Select = gSL_Select & " AND   値 = 'PNJ0310'"
    If SYMRead(gSL_Select, 1) Then
       iH1_YM = RTrim(SYM.英数字1)
       iH1_NYM = RTrim(SYM.英数字1)         '2010/11/11
    Else
       'iH1_YM = ""                         '2010/11/11
       iH1_YM = Format(Now, "YYYYMM")       '2010/11/11
       iH1_NYM = Format(Now, "YYYYMM")      '2010/11/11
    End If

    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

'   CSVﾌｧｲﾙ出力先
    mCsvName = Mid(CisFun.WinRoot, 1, 3) & "工事手配書.csv"
    mCsvKey = "CNJ1130"
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
        .Gradation 50, 130, 230, 4, 2, 3
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
    If Dummy.Tag = "iH1_Torcd" Then Call PB_Torcd_Click: Exit Sub
End Sub
'+---------------------+
'+    取引先検索        +
'+---------------------+
Private Sub PB_Torcd_Click()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
'
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+    iH1_Torcd(GotFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+----------------------------+
'+    iH1_Torcd(LostFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_LostFocus()
    Dummy.Tag = ""
    H1lb_Tornm = ""
    If RTrim(iH1_Torcd) <> "" Then
        If TorNmGet(iH1_Torcd, 8) Then
            H1lb_Tornm = TRM_RName
        End If
    End If
End Sub
'+----------------------------------------+
'+  Op_Out   出力先
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

'   年月
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    [年月]を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    If Not iH1_YM.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    [年月]を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If

'================================ 2010/11/11 Start
'  計画年月
    If Trim(iH1_NYM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    [計画年月]を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_NYM.SetFocus
        Exit Function
    End If
    If Not iH1_NYM.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    [計画年月]を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_NYM.SetFocus
        Exit Function
    End If
'================================ 2010/11/11 End

'   仕入先
    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 8) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    仕入先　未登録！（取引先マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If

    lb_msg.Caption = "【 発注内示データ 抽出中 】"
    DoEvents

'======================================================================== 2010/11/11 Start
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "SELECT"
'    gSL_Select = gSL_Select & "  HN.手配先,ISNULL(TR.略称,'') 手配先名,ISNULL(TR.工場名,'') 機械,"    '!機械項目間借り
'    gSL_Select = gSL_Select & "  HN.品番,ISNULL(HM.表示品番,'') 品番編集, ISNULL(HM.品名,'') 品名,"
'    gSL_Select = gSL_Select & "  SUM( HN.数量 ) 数量"
'    gSL_Select = gSL_Select & " FROM"
'    gSL_Select = gSL_Select & "  発注内示テーブル HN"
'    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 取引先マスタ TR"
'    gSL_Select = gSL_Select & "    ON TR.取引先CD = HN.手配先"
'    gSL_Select = gSL_Select & "   AND TR.取引先区分 = 1"
'    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 品番マスタ HM"
'    gSL_Select = gSL_Select & "    ON HN.品番 = HM.品番"
'    gSL_Select = gSL_Select & " WHERE HN.内示年月 = '" & RTrim(iH1_YM) & "'"
'    gSL_Select = gSL_Select & "   AND HN.計画年月 = '" & RTrim(iH1_YM) & "'"
'    gSL_Select = gSL_Select & "   AND ISNULL(TR.社内区分,0) = 0 "
''   >> 手配先
'    If RTrim(iH1_Torcd) <> "" Then
'        gSL_Select = gSL_Select & "   AND HN.手配先 = '" & RTrim(iH1_Torcd) & "'"
'    End If
'    gSL_Select = gSL_Select & " GROUP BY HN.手配先,ISNULL(TR.略称,''),ISNULL(TR.工場名,''),"
'    gSL_Select = gSL_Select & " HN.品番,ISNULL(HM.表示品番,''), ISNULL(HM.品名,'')"
'    gSL_Select = gSL_Select & " ORDER BY "
'    gSL_Select = gSL_Select & "  HN.手配先,HN.品番"
'


    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT"
    gSL_Select = gSL_Select & "  HT.取引先 手配先,ISNULL(TR.略称,'') 手配先名,ISNULL(TR.工場名,'') 機械,"    '!機械項目間借り
    gSL_Select = gSL_Select & "  HT.品番,ISNULL(HM.表示品番,'') 品番編集, ISNULL(HM.品名,'') 品名,"
    gSL_Select = gSL_Select & "  SUM( ISNULL(HN.数量,0) ) 数量"
    gSL_Select = gSL_Select & " FROM"
    gSL_Select = gSL_Select & " 品番取引先マスタ HT "
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 発注内示テーブル HN"
    gSL_Select = gSL_Select & "    ON HN.内示年月 = '" & RTrim(iH1_YM) & "'"
'    gSL_Select = gSL_Select & "   AND HN.計画年月 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   AND HN.計画年月 = '" & RTrim(iH1_NYM) & "'"
    gSL_Select = gSL_Select & "   AND HN.手配先 = HT.取引先 "
    gSL_Select = gSL_Select & "   AND HN.受入 = HT.受入 "
    gSL_Select = gSL_Select & "   AND HN.品番 = HT.品番 "
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "    ON TR.取引先CD = HT.取引先"
    gSL_Select = gSL_Select & "   AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 品番マスタ HM"
    gSL_Select = gSL_Select & "    ON HT.品番 = HM.品番"
    gSL_Select = gSL_Select & "  WHERE ISNULL(TR.社内区分,0) = 0 "
    gSL_Select = gSL_Select & "    AND HT.品目 = 1 "
'   >> 手配先
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "   AND HT.取引先 = '" & RTrim(iH1_Torcd) & "'"
    End If
    gSL_Select = gSL_Select & " GROUP BY HT.取引先,ISNULL(TR.略称,''),ISNULL(TR.工場名,''),"
    gSL_Select = gSL_Select & " HT.品番,ISNULL(HM.表示品番,''), ISNULL(HM.品名,'')"
    gSL_Select = gSL_Select & " ORDER BY "
    gSL_Select = gSL_Select & "  HT.取引先,HT.品番"
'======================================================================== 2010/11/11 End


    If Not HNJRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

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

'  ワークテーブル　削除
    With CisDB
        '+---------- ワークの削除 -----------+
        .SQL = ""
        .SQL = .SQL & "  DELETE FROM 工事手配書ワーク"
        .SQL = .SQL & "  WHERE 端末番号 = SUBSTRING(HOST_NAME(), 1, 20)"
        Call CisDB.DBExec
    End With
    
    'ワークファイル出力
    If Not Work_Out Then Exit Function
    'ワークファイル読み込み
    gSL_Select = "SELECT * FROM 工事手配書ワーク "
    gSL_Select = gSL_Select & " WHERE 端末番号 = SUBSTRING(Host_Name(),1,20) "
'    gSL_Select = gSL_Select & " AND   ISNULL(受注実績見込,0) <> 0 "               '2010/11/11          2012/11/26 中止
    gSL_Select = gSL_Select & " AND   ( ISNULL(受注実績見込,0) <> 0 OR ISNULL(当月所要量,0) <> 0 ) "   '2012/11/26
    gSL_Select = gSL_Select & " ORDER BY 手配先,品番 "                            '2010/11/11
    Call KJTRead(gSL_Select)
    
    If Op_Out(0).Value Or Op_Out(1).Value Then
        lb_msg.Caption = "【 工事手配データ 印刷中 】"
        DoEvents
        
        With CisPrtCNJ1130
            .cDataMaxCnt = CisDB.RecordCount
            Set .iPrtForm = CNJ1130L
            .dpMSG = "工事手配書　印刷中"
            If Op_Out(0).Value Then
                .iPreview = pv_Print
            Else
                .iPreview = pv_Preview
            End If
            
            DoEvents
            
            .PrintStart
        End With
    Else
        If Not CsvPut Then Exit Function
    End If
    PrintStart = True
End Function
Private Function Work_Out() As Boolean
    Dim wNounyuSu       As Long

    Work_Out = False
    
    '前月取得
    With CisFun
        .Date_RsFormat = fYM
'        .Date_Add iH1_YM, M, -1                    '2010/11/11
        .Date_Add iH1_NYM, M, -1                    '2010/11/11
        wZYM = Mid$(.Date_Result, 1, 6)
    End With

' ワークテーブル出力(親品番取得、親品番情報取得)
    Do Until Not HNJ_RDSTS
    
        wNounyuSu = 0
        
        With CisDB
            .SQL = "構成問合せ情報抽出"
            .StoadoCount = 4
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = Input用
            .ParaNo = 3: .ParaIO = Input用
            .ParaNo = 4: .ParaIO = Input用
        
            '逆展開/製品 検索
            .ParaNo = 1: .ParaValue = RTrim(HNJ.品番)
            .ParaNo = 2: .ParaValue = 0
            .ParaNo = 3: .ParaValue = 1
            .ParaNo = 4: .ParaValue = 1
            
        
            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     構成マスタ情報の抽出に失敗しました。     "
                    .MB_Title = "抽出処理"
                    .MB_Button = Error
                    If .MBOX Then Exit Function
                End With
            End If
            
            gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示品番 FROM 構成問合せワーク KW "
            gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
            gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示品番"
            gSL_Select = gSL_Select & " ORDER BY KW.親品番"
            
            If KOLRead(gSL_Select, , 1) Then
               Do Until Not KOL_RDSTS
                  gSL_Select = "SELECT ISNULL(SUM(納入数),0) 納入数 FROM 受注テーブル "
                  gSL_Select = gSL_Select & " WHERE SUBSTRING(納入日,1,6) = '" & wZYM & "'"
                  gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(KOL.親品番) & "'"
                  If JYTRead(gSL_Select, 1, 2) Then
                     wNounyuSu = wNounyuSu + JYT.納入数
                  End If
                  Call KOLReadNext(1)
               Loop
            Else
                '==================================- 2011/11/28 Start
                gSL_Select = "SELECT ISNULL(SUM(納入数),0) 納入数 FROM 受注テーブル "
                gSL_Select = gSL_Select & " WHERE SUBSTRING(納入日,1,6) = '" & wZYM & "'"
                gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(HNJ.品番) & "'"
                If JYTRead(gSL_Select, 1, 2) Then
                   wNounyuSu = wNounyuSu + JYT.納入数
                End If
                '==================================- 2011/11/28 End
            End If
            Call KOLClose(1)
            
        End With
                    
        '工事手配書ワーク出力
        With KJT
            .端末番号 = CisFun.Mid2(CisFun.WSID, 1, 20)
            .手配先 = RTrim(HNJ.手配先)
            .手配先名 = RTrim(HNJ.手配先名)
            .工場名 = RTrim(HNJ.機械)               '機械間借り
            .品番 = RTrim(HNJ.品番)
            .表示品番 = RTrim(HNJ.品番編集)
            .品名 = RTrim(HNJ.品名)
            .受注実績見込 = wNounyuSu
            .当月所要量 = HNJ.数量
        End With
        Call KJTInsert(1)
        
        
        Call HNJReadNext(0)
    Loop
    
    Call HNJClose
    
    Work_Out = True
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

    lb_msg.Caption = "【 工事手配書データ 出力中 】"
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
    Do Until Not KJT_RDSTS
    '   >> ｶｳﾝﾄ
        OutCnt = OutCnt + 1
    '   >> Body部 編集＆出力
        Call CsvEdit_Body1(OutStr, OutCnt)
        Print #OutFNo, OutStr

        CisFrm.MeterUpdate OutCnt

    '   >> 次のﾚｺｰﾄﾞを読み込む
        Call KJTReadNext
    Loop
    Call HNJClose
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
    CsvStr = CsvStr & "工事手配書"
End Sub
Private Sub CsvEdit_Head2(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & "№"
    CsvStr = CsvStr & ",年月"
    CsvStr = CsvStr & ",仕入先"
    CsvStr = CsvStr & ",仕入先名"
    CsvStr = CsvStr & ",工場名"
    CsvStr = CsvStr & ",品番"
    CsvStr = CsvStr & ",品名"
    CsvStr = CsvStr & ",受注実績見込"
    CsvStr = CsvStr & ",当月所要量"
End Sub
'+----------------------------+
'+    Body部　編集処理
'+----------------------------+
Private Sub CsvEdit_Body1(CsvStr As String, CsvCnt As Long)
    With KJT
        CsvStr = ""
        CsvStr = CsvStr & CsvCnt
'        CsvStr = CsvStr & "," & RTrim(iH1_YM)                  '2010/11/11
        CsvStr = CsvStr & "," & RTrim(iH1_NYM)                  '2010/11/11
        CsvStr = CsvStr & "," & RTrim(.手配先)
        CsvStr = CsvStr & "," & RTrim(.手配先名)
        CsvStr = CsvStr & "," & RTrim(.工場名)
        CsvStr = CsvStr & "," & RTrim(.表示品番)
        CsvStr = CsvStr & "," & RTrim(.品名)
        CsvStr = CsvStr & "," & .受注実績見込
        CsvStr = CsvStr & "," & .当月所要量
    End With
End Sub




