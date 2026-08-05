VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CisYM_6.0.ocx"
Begin VB.Form CTA0510 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "棚卸HHT読取結果取込"
   ClientHeight    =   11190
   ClientLeft      =   3675
   ClientTop       =   1875
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   Begin VB.PictureBox Bar 
      Height          =   555
      Left            =   4800
      ScaleHeight     =   495
      ScaleWidth      =   5715
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   5775
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5520
      Top             =   7800
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
   Begin VB.Timer Timer2 
      Left            =   945
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Left            =   510
      Top             =   840
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   840
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8385
      Top             =   7800
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
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 棚卸HHT読取結果取込 】"
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
         Left            =   14370
         Top             =   60
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
   Begin Cis3D_v60.CIS3D PB_Ok 
      Height          =   585
      Left            =   6960
      Top             =   7800
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
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   3820
      Left            =   4800
      ScaleHeight     =   3918.678
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   5700
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3075
      Width           =   5760
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   3675
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   6482
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
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   420
            Left            =   75
            Top             =   2220
            Width           =   5430
            _ExtentX        =   9578
            _ExtentY        =   741
            ForeColor       =   16711680
            Caption         =   "今回対象年月"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            Begin CisYM_V60.CisYM CisYM1 
               Height          =   330
               Left            =   1785
               TabIndex        =   14
               Top             =   60
               Width           =   1320
               _ExtentX        =   2328
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
               cChkResult      =   0   'False
            End
            Begin Cis3D_v60.CIS3D lb_RunCount 
               Height          =   405
               Left            =   3555
               Top             =   0
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   714
               BackColor       =   16761024
               Caption         =   "ZZZ / ZZZ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
               cAlingnment     =   4
               cPositionX      =   -500
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   945
            Left            =   60
            Top             =   2670
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   1667
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
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   885
               Left            =   30
               Top             =   30
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   1561
               Caption         =   "今回処理対象ファイル数"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D lb_FSu 
               Height          =   885
               Left            =   3555
               Top             =   30
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   1561
               Caption         =   "ZZZ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   27.75
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
               cAlingnment     =   4
               cPositionX      =   -500
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D28 
            Height          =   615
            Index           =   1
            Left            =   60
            Top             =   60
            Width           =   5430
            _ExtentX        =   9578
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
            Begin VB.PictureBox B1_Area1 
               BackColor       =   &H00FFC0FF&
               Height          =   555
               Left            =   30
               ScaleHeight     =   495
               ScaleWidth      =   5310
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   30
               Width           =   5370
               Begin Cis3D_v60.CIS3D lb_msg 
                  Height          =   375
                  Left            =   -240
                  Top             =   90
                  Width           =   5790
                  _ExtentX        =   10213
                  _ExtentY        =   661
                  BackColor       =   16761087
                  Caption         =   "棚卸HHT読取結果取込処理を行います"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   1395
            Left            =   60
            Top             =   780
            Width           =   5430
            _ExtentX        =   9578
            _ExtentY        =   2461
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
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "前回年月  ："
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   240
               Left            =   510
               TabIndex        =   13
               Top             =   105
               Width           =   1440
            End
            Begin VB.Label Hlb_ZYM 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "YYYY年MM月"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   240
               Left            =   2040
               TabIndex        =   12
               Top             =   120
               Width           =   1200
            End
            Begin VB.Label Hlb_FSu 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "Z,ZZZ"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   240
               Left            =   2040
               TabIndex        =   10
               Top             =   555
               Width           =   600
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "前回ﾌｧｲﾙ数："
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   240
               Left            =   510
               TabIndex        =   9
               Top             =   540
               Width           =   1440
            End
            Begin VB.Label Hlb_ZDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "9999年99月99日 99：99"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   240
               Left            =   2040
               TabIndex        =   8
               Top             =   990
               Width           =   2520
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "前回処理日："
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   240
               Left            =   510
               TabIndex        =   7
               Top             =   990
               Width           =   1440
            End
         End
      End
   End
   Begin VB.PictureBox CrDraw1 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label PNL_Ok 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6945
      TabIndex        =   6
      Top             =   8475
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8385
      TabIndex        =   3
      Top             =   8475
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5505
      TabIndex        =   2
      Top             =   8475
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CTA0510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  ＨＨＴ読取結果取込
'**       フォームID    :  CTA0510              （Based CXJ0710 テキスト受注データ変換）
'**
'**       処理概要      :
'**
'**
'**       作  成  日    :   2019/07/03
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

    Dim ZYM             As String           ' 2019/07/04 ADD
    Dim ZDate           As String
    Dim ZFSu            As Long
    Dim ZProc           As String
    
    Dim mKaraPrint      As Byte
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/03 INSERT START
    Dim iMovFile        As Long             ' 退避する.CSVファイル数
    Dim wMovFile()      As String           ' 退避する.CSVファイル名
    Dim wCsvFLD_Get     As String           ' .CSVファイル取得フォルダ
    Dim wCsvFLD_Mov     As String           ' .CSVファイル退避フォルダ
    Dim iCsvFileCnt     As Long
    Dim wStrFName       As String
    Dim iNoHinban       As Long
    
    Const cBackFLD      As String = "Back"
    Public Text_Frm     As CISFormContorl
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/03 INSERT START
'
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止 #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # 初 期 内 容 取 得 #
'   #-------------------#
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/11/15 INSERT START
    CisDB.TimeOut = 3600
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/11/15 INSERT END
    If Not IniGet Then End

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/03 INSERT START
    ' .CSVファイルのフォルダ情報取得
    With CisFun
        .INI_Section = "棚卸読取結果"
        .INI_FileID = gIniExe & gIniName
        ' サーバー名
        .INI_Default = ""
        .INI_Key = "FLD"
        .GetIni
        wCsvFLD_Get = RTrim(.INI_String)
        If wCsvFLD_Get = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   棚卸読取結果フォルダ情報がありません。          "
                .MB_MSG(3) = "       処理を終了します。                          "
                .MB_Title = ".iniファイルエラー"
                .MB_Button = OK
                Call .MBOX
            End With
            End
        End If
        wCsvFLD_Mov = wCsvFLD_Get & cBackFLD
    End With
   
   
'    ' ※※※※※※※※※※
'    Dim objFSO As Object
'    Set objFSO = CreateObject("Scripting.FileSystemObject")
'
''    Dim aa As Integer
''    With objFSO.GetFolder(wCsvFLD_Get)
''        aa = .Files.Count
''    End With
'
''    iCsvFileCnt = 0
''    wStrFName = Dir(wCsvFLD_Get & "*.csv")      ' 最初のファイル名を返します。
''    Do While wStrFName <> ""
''        ''''Call MsgBox(wStrFName, vbOKOnly)    ' 確認用
''        iCsvFileCnt = iCsvFileCnt + 1
''        wStrFName = Dir                         ' 次のファイル名を返します。
''    Loop
''    If iCsvFileCnt = 0 Then
'    If CSV_GetFileCount = 0 Then
''        Call MsgBox("棚卸読取結果の.CSVファイルがありません" & vbCrLf + vbCrLf & _
''                    "　　　　　処理を終了します　　　　　　", vbOKOnly, "ファイルエラー")
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "   棚卸読取結果の.CSVファイルがありません。        "
'            .MB_MSG(3) = "       処理を終了します。                          "
'            .MB_Title = "処理ファイルエラー"
'            .MB_Button = OK
'            Call .MBOX
'        End With
'        End
'    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/03 INSERT START
    
    mSaveMsg = lb_msg
'   #-------------------#
'   #  項 目 初 期 化   #
'   #-------------------#
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
        .MSGType = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

'   #-------------------#
'   #    初期入力設定   #
'   #-------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 対象ファイル取得  #
'   #-------------------#
    gLong = CSV_GetFileCount
    lb_FSu = Format(gLong, "#0")
'   #-------------------#
'   # 画面中央位置表示  #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    If gLogoDisp Then
       BG_Logo(0).Caption = gLogo
       CisFrm.BGLDisp BG_Logo, 10
    End If
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/03 INSERT START
'    Hlb_ZYM.Caption = ""
'    Hlb_FSu.Caption = ""
'    Hlb_ZDate.Caption = ""
'    Call ItemsClearSYM
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "SELECT ISNULL(英数字1, '') 英数字1, " + vbCrLf
'    gSL_Select = gSL_Select & "       ISNULL(数字1  ,  0) 数字1  , " + vbCrLf
'    gSL_Select = gSL_Select & "       ISNULL(英数字2, '') 英数字2  " + vbCrLf
'    gSL_Select = gSL_Select & "  FROM 名称マスタ                   " + vbCrLf
'    gSL_Select = gSL_Select & " WHERE 区分名称 = '棚卸読取結果取込'" + vbCrLf
'    gSL_Select = gSL_Select & "   AND 値       = '1'               " + vbCrLf
'    Call SYMRead(gSL_Select, 1)
'    If SYM_RDSTS Then
'        Hlb_ZYM.Caption = Mid(SYM.英数字1, 1, 4) & "年" & Mid(SYM.英数字1, 5, 2) & "月"
'        Hlb_FSu.Caption = Format(SYM.数字1, "#,##0")
'        Hlb_ZDate.Caption = Mid(SYM.英数字2, 1, 4) & "年" & Mid(SYM.英数字2, 5, 2) & "月" & _
'                            Mid(SYM.英数字2, 7, 2) & "日" & " " & _
'                            Mid(SYM.英数字2, 9, 2) & ":" & Mid(SYM.英数字2, 11, 2)
'    End If
    
    CisYM1.Text = Format(Now(), "yyyyMM")
    lb_RunCount.Visible = False
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/03 INSERT END
    
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 INSERT START
    ' ﾌﾟﾘﾝﾀ出力先
    mPrtDev = Printer.DeviceName
    mSavePrint = mPrtDev
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 INSERT END
    
    FormAct = Not FormAct
    
'< 前回処理日 >
    Call ZenProc_Get
    Hlb_ZDate = RTrim$(ZDate)
    If Trim(ZDate) <> "" Then
       Hlb_ZDate = Hlb_ZDate & ZProc
    End If
    Hlb_FSu = Format(ZFSu, "###")
    Hlb_ZYM = ZYM                       ' 2019/07/04 ADD
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
        Case vbKeyF12
                            If PB_END.Visible Then Call PB_END_Click        '【終了】
                            If PB_Ok.Visible Then Call PB_END_Click         '【終了】
        Case vbKeyReturn
                            If PB_Ent.Visible Then                          '【入力】
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
       '.Gradation 100, 160, 250, 4, 2, 3
        .Gradation &H0, &H0, &H7F, 3, 3, 3 '0, 0, 10, 3, 3, 3
    End With
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
  ''Call SingleProcess_Release  ' 2019/07/05 DEL
    Call End_Proc
End Sub
Private Sub End_Proc()
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp
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
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   Call CisFrm.TimeSet("1")
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
    Dim B1Mode      As Boolean
    Dim B2Mode      As Boolean

    H1Mode = False
    B1Mode = False
    B2Mode = False
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True
       Case "B1":   B1Mode = True
       Case "B2":   B2Mode = True
    End Select
    H1_Area1.Enabled = H1Mode
  ''B1_Area1.Visible = B1Mode Or B2Mode     ' 2019/07/08 DEL    ' 2019/07/05 UPD：B2Modeを条件に追加
    B1_Area1.Visible = Not B1Mode           ' 2019/07/08 ADD

    PB_Ent.Visible = H1Mode
    PNL_ENT.Visible = H1Mode
    PB_END.Visible = H1Mode
    PNL_END.Visible = H1Mode
    Timer2.Enabled = H1Mode Or B2Mode
    PB_Ok.Visible = B2Mode
    PNL_Ok.Visible = B2Mode


' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    Static sw_Str           As String
'
    gLong = CSV_GetFileCount
    lb_FSu = Format(gLong, "#0")

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/04 INSERT START
    If (Not CisYM1.cChkResult) Or _
        Trim(CisYM1.Text) = "" Then
           With CisFun
                .MB_Lines = 3
                .MB_MSG(2) = "     今回対象年月を正しく入力してください。        "
                .MB_Title = "年月エラー"
                .MB_Button = OK
                Call .MBOX
           End With
           GoTo ReturnPress_Ee
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/04 INSERT END
    
    If CisFun.Val2(gLong) = 0 Then
       With CisFun
            .MB_Lines = 3
            .MB_MSG(2) = "       処理対象ファイルがありません。              "
            .MB_Title = "テキスト変換処理"
            .MB_Button = OK
            Call .MBOX
       End With
       GoTo ReturnPress_Ee
    End If
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "       処理を行います。         "
        .MB_MSG(4) = "       よろしいですか？         "
        .MB_Title = "テキスト受注ﾃﾞｰﾀ変換処理"
        .MB_Button = Yes_No
        If Not .MBOX Then GoTo ReturnPress_Ee
    End With
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/04 INSERT START
    ProcHB = "B1"
    Call DispChange(ProcHB)

    With CisDB
        .SQL = ""
        .SQL = .SQL + "Delete From 棚卸読取結果テーブル" + vbCrLf
        .SQL = .SQL + " Where 対象年月 = '" + CisYM1.Text + "'"
        .DBExec
    End With
    
    Call SaveDB
    
    bErrorHak = False
    If iNoHinban <> 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "       品番マスタ未登録データが " & CStr(iNoHinban) & "件 あります。         "
            .MB_MSG(4) = "       エラーリストを発行しますか？               "
            .MB_Title = "エラーリスト発行確認"
            .MB_Button = Yes_No
            bErrorHak = .MBOX
        End With
    End If
    
    If bErrorHak Then
      ''MsgBox ("エラーリスト発行しますよ")
        PrintStart
    End If
    
    Call ZenProc_Put

    ProcHB = "B2"
    Call DispChange(ProcHB)
    GoTo ReturnPress_Ed
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/04 INSERT END




    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE START
'    Timer2.Enabled = False
'    DoEvents
'
'    On Error GoTo ReturnPress_Er
'
''------------------- 履歴処理開始
''
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE END
ReturnPress_Ee:
'------------------------------------- 再入力
ReturnPress_Ed:
    gStr = "【  処　理  終  了  】"     ''"棚卸HHT読取結果取込処理　終了"
    lb_msg.Caption = gStr
    Timer2.Enabled = True
    With CisFrm
        .MSGText = lb_msg.Caption
        .MSGType = FSize_Changes
        .MSGInterval = 100
        .MSGFSizeMax = 18
        .MSGRelate
    End With
    ProcHB = "B2"
    Call DispChange(ProcHB)
    Call CisFrm.MousePT(1)
End Sub

'*-------------------------------------------------*
'*              前　回　処　理　日      ( 獲得 )
'*-------------------------------------------------*
Private Function ZenProc_Get()
    ZDate = ""
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "棚卸読取結果取込"          ' 2019/07/04 UPD："テキスト管理[愛知機械]"      ' 2009.09.15 upd : "新eかんばん管理"
        .区分桁数 = 1                           ' 2019/07/04 UPD：2
        .値 = "1"                               ' 2019/07/04 UPD："JY"
        .制御区分 = "X"
        .値名称 = "前回内容"                    ' 2019/07/04 UPD："テキスト受注[愛知機械]"      ' 2009.09.15 upd : "新eかんばん受注"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)
    
    Call ItemsClearSYM
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称   = '棚卸読取結果取込'"      ' 2019/07/04 UPD：'テキスト管理[愛知機械]'"    ' 2009.09.15 upd : '新eかんばん管理'
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値         = '1'"                     ' 2019/07/04 UPD：'JY'"
    Call SYMRead(gSL_Select, 1)
    
    ZDate = RTrim$(SYM.英数字2)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/04 INSERT START
    ZProc = ""
    ZYM = ""
    If Trim(SYM.英数字1) <> "" Then
        ZYM = Mid(SYM.英数字1, 1, 4) & "年" & _
              Mid(SYM.英数字1, 5, 2) & "月"
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/04 INSERT START
    ZFSu = SYM.数字1                                                ' 2019/07/04 UPD：SYM.数字2
End Function
'*-------------------------------------------------*
'*              前　回　処　理　日      ( 書込 )
'*-------------------------------------------------*
Private Function ZenProc_Put()

    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = '" & CisYM1.Text & "',"           ' 2019/07/04 UPD：'',"
    gSL_Select = gSL_Select & " 英数字2 = '" & Format(Now(), "yyyy年mm月dd日 HH:MM") & "',"
    gSL_Select = gSL_Select & " 数字1   =  " & iCsvFileCnt & ","            ' 2019/07/04 UPD：Text_ProcKB
    gSL_Select = gSL_Select & " 数字2   = 0,"                               ' 2019/07/04 UPD：" & CisFun.Val2(lb_FSu) & ","
    gSL_Select = gSL_Select & " 数字3   = 0,"                               ' 2019/07/04 UPD：" & Text_Rireki & ","
    gSL_Select = gSL_Select & " 更新日  = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者  = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末= SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = '棚卸読取結果取込'"        ' 2019/07/04 UPD：'テキスト管理[愛知機械]'"  ' 2009.09.15 upd : '新eかんばん管理'
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '1'"                             ' 2019/07/04 UPD：'JY'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

End Function




'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/03 INSERT START
'+------------------------------------------------------------------------+
'+
'+      ﾃｷｽﾄ連携格納ファイル数を求める
'+
'+------------------------------------------------------------------------+
Public Function CSV_GetFileCount() As Long
    Dim wCount      As Long
    Dim wStrFName   As String
    wCount = 0
    
    On Error GoTo CSV_GetFileCount_er:

'    If Text_Rireki = 0 Then
'        wStrFName = Dir(Text_InDir(Text_Proc) & Text_TextName(Text_Proc))      ' 最初のファイル名を返します。
'    Else
'        wStrFName = Dir(TRI.処理退避先 & Text_TextName(Text_Proc))      ' 最初のファイル名を返します。
'    End If
    wStrFName = Dir(wCsvFLD_Get & "*.csv")      ' 最初のファイル名を返します。
    Do While wStrFName <> ""
        wCount = wCount + 1
        wStrFName = Dir                         ' 次のファイル名を返します。
    Loop
    CSV_GetFileCount = wCount
    Exit Function
CSV_GetFileCount_er:
    CSV_GetFileCount = 0
End Function







Private Function SaveDB() As Boolean
    SaveDB = False
    
    On Error GoTo SaveDB_Err
    
    Dim wCount      As Long
    Dim wStrFName   As String
    
    iNoHinban = 0                               ' 品番なし数初期化
    Call Insert_ErrorData(iNoHinban, "")        ' 品番無しエラーテーブル初期化
    
    iMovFile = 0
    wCount = 0
    wStrFName = Dir(wCsvFLD_Get & "*.csv")      ' 最初のファイル名を返します。
    Do While wStrFName <> ""
        lb_RunCount.Caption = CStr(wCount + 1) + " / " + lb_FSu.Caption
        lb_RunCount.Visible = True
        DoEvents
        
        Call CSVtoDB(wCsvFLD_Get & wStrFName)
                
      ''Move_CSV (wStrFName)
        iMovFile = iMovFile + 1                 ' 移動するファイル名を保存
        ReDim Preserve wMovFile(iMovFile)
        wMovFile(iMovFile) = wStrFName
        
        wCount = wCount + 1
        wStrFName = Dir                         ' 次のファイル名を返します。
    Loop
    
    lb_RunCount.Caption = ""
    lb_RunCount.Visible = False
    
    Call Move_CSV
    
    iCsvFileCnt = wCount
    SaveDB = True
    Exit Function

SaveDB_Err:
    On Error GoTo 0
    SaveDB = False
End Function






Private Sub CSVtoDB(ByVal wFName As String)
    'CSV ファイルを読み込み MSFlexGrid に表示
    Dim intFileNo   As Integer      'ファイルNo
    Dim TextLine    As String
    Dim CellsData   As Variant
    Dim i           As Long
    
    Dim wYM         As String
    Dim lCt         As Long
    Dim fDB         As Boolean
    
    wYM = CisYM1.Text
    lCt = 0
    
    intFileNo = FreeFile
    
    
On Error GoTo err1
    
    
    ' 進行状況バーの設定
    Dim kensu As Long
    kensu = CreateObject("Scripting.FileSystemObject"). _
                    OpenTextFile(wFName, 8).Line - 1
    With CisFrm
        Set .MeterPic = Bar
        .MeterMax = CisFun.Val2(kensu)
    End With
    Bar.Visible = True
    Set Text_Frm = CisFrm
    
    
    Open wFName For Input As #intFileNo
    Do Until EOF(intFileNo)                             'EOF(intFileNo)が True になるまで実行
        Line Input #intFileNo, TextLine                 '1行全体を変数に読み込む
'        '文字列中の "" も取り除かれるので注意(別途追加設定が必要)
'        TextLine = Replace(TextLine, Chr$(34), "")      ' "" を取り除く
        CellsData = Split(TextLine, ",")                'カンマ区切りで列データを分割
'        For i = LBound(CellsData) To UBound(CellsData)
'            Text1.Text = Text1.Text & CellsData(i) & vbTab
'        Next i
'        Text1.Text = Text1.Text & vbCrLf
    
        fDB = True
        With CisDB
            .SQL = ""
            .SQL = .SQL + "SELECT 品番 FROM 品番マスタ" & vbCrLf
            .SQL = .SQL + " WHERE 品番 = '" & RTrim(CStr(CellsData(7))) & "'" & vbCrLf
            .DBExec
            If Not .DBRead Then
                fDB = False
                iNoHinban = iNoHinban + 1
                Call Insert_ErrorData(iNoHinban, TextLine)      ' 印刷処理のデータ作成
            End If
        End With
    
        If fDB Then
            Call ItemsClearHHR
            With HHR
                .対象年月 = wYM
                .処理日 = CisFun.Mid2(RTrim(CStr(CellsData(0))), 1, 10)
                .処理時間 = CisFun.Mid2(RTrim(CStr(CellsData(1))), 1, 8)
                .機種名 = CisFun.Mid2(RTrim(CStr(CellsData(2))), 1, 30)
                .担当者 = CisFun.Mid2(RTrim(CStr(CellsData(3))), 1, 20)
                .置場 = CisFun.Mid2(RTrim(CStr(CellsData(4))), 1, 30)
                .読取連番 = CLng(CellsData(5))
                .かんばんQR内容 = CisFun.Mid2(RTrim(CStr(CellsData(6))), 1, 300)
                .品番 = CisFun.Mid2(RTrim(CStr(CellsData(7))), 1, 30)
                .実棚数 = CLng(CellsData(8))
                .作成者 = gTanto
            End With
            Call HHRInsert
        End If
            
        ' 進行状況表示
        lCt = lCt + 1
        If Not Text_Frm Is Nothing Then
            Text_Frm.MeterUpdate lCt
        End If

    Loop
    
GoTo CSVtoDB_Ed
    
err1:
    Dim a As String
    a = "ERR"
CSVtoDB_Ed:
    Close #intFileNo
    Bar.Visible = False
End Sub


'+---------------------------------------------+
'+      印刷処理                               +
'+---------------------------------------------+
Private Function PrintStart() As Boolean
    PrintStart = False
        lb_msg.Caption = "【 エラーリスト 印刷中 】"
        DoEvents
        With CisPrtCTA0510
            .cDataMaxCnt = iNoHinban 'CisDB.RecordCount
            Set .iPrtForm = CTA0510L
            .dpMSG = "エラーリスト 印刷中"
'            If Op_Out(0).Value Then
                .iPreview = pv_Print
'            Else
'                .iPreview = pv_Preview
'            End If
            .PrintStart
        End With
    PrintStart = True
End Function

'+---------------------------------------------+
'+      CSV 退避                               +
'+---------------------------------------------+
Private Function Move_CSV() As Boolean
    Move_CSV = False
    
    If iMovFile = 0 Then
        Move_CSV = True
        Exit Function
    End If
    
    Dim iLoop   As Integer
    Dim wFName  As String
    
    Dim objFSO  As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
        
    ' Backフォルダ存在確認：無ければ作成する
    If Not objFSO.folderexists(wCsvFLD_Mov) Then
        CisFun.FE_Path = wCsvFLD_Mov
        CisFun.DirCheck NoMsg_Crt
    End If
        
    ' 当日フォルダ存在確認：無ければ作成する
    wFName = wCsvFLD_Mov & "\" & Format(Now(), "yyyyMMdd")
    If Not objFSO.folderexists(wFName) Then
        CisFun.FE_Path = wFName
        CisFun.DirCheck NoMsg_Crt
    End If
        
    ' 連番フォルダ存在確認：最大値のフォルダが無ければ作成する
    For iLoop = 1 To 999
       'Dim aaa As String: aaa = Format(iLoop, "000")   ''確認用
        If Not objFSO.folderexists(wFName & "\" & Format(iLoop, "000")) Then
            wFName = wFName & "\" & Format(iLoop, "000")
            CisFun.FE_Path = wFName
            CisFun.DirCheck NoMsg_Crt
            Exit For
        End If
    Next
    If iLoop > 999 Then
      ''MsgBox ("999個 超えたよ")
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    退避フォルダの数が上限に達しました。            "
            .MB_MSG(3) = "    現在処理された.CSVファイルは退避されません。    "
            .MB_Title = "フォルダ作成エラー"
            .MB_Button = OK
            Call .MBOX
        End With
        Exit Function
    End If
    
    ' 移動
    For iLoop = 1 To iMovFile
        FileCopy wCsvFLD_Get + "\" + wMovFile(iLoop), wFName + "\" + wMovFile(iLoop)
        Kill wCsvFLD_Get + "\" + wMovFile(iLoop)
    Next
    
    Move_CSV = True
    
Move_CSV_Ed:
  ''objFSO = Nothing        ' これ実施するとエラーになる
End Function
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/03 INSERT END

