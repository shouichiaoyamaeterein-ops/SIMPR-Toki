VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{A37A80A5-3F34-11D2-912B-00C0DF49BE5E}#2.0#0"; "CrDraw.ocx"
Begin VB.Form CXJ0710 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "【愛知機械】受注データ変換"
   ClientHeight    =   11190
   ClientLeft      =   150
   ClientTop       =   1290
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
      TabIndex        =   12
      Top             =   3450
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
      Top             =   855
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
      Caption         =   "【 [愛知機械]受注データ変換 】"
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
      Height          =   3015
      Left            =   4800
      ScaleHeight     =   3075.616
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   5700
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4050
      Width           =   5760
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   2850
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   5027
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   945
            Left            =   60
            Top             =   1860
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
            Begin VB.PictureBox B1_Area 
               BackColor       =   &H00FFC0FF&
               Height          =   555
               Left            =   45
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
                  Caption         =   "[愛知機械]受注データ変換処理を行います"
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
            Height          =   1035
            Left            =   60
            Top             =   780
            Width           =   5430
            _ExtentX        =   9578
            _ExtentY        =   1826
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
               TabIndex        =   11
               Top             =   180
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
               TabIndex        =   10
               Top             =   150
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
               TabIndex        =   9
               Top             =   600
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
               TabIndex        =   8
               Top             =   600
               Width           =   1440
            End
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   1125
      Left            =   4950
      ScaleHeight     =   1065
      ScaleWidth      =   5385
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5790
      Visible         =   0   'False
      Width           =   5445
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   525
         Left            =   0
         Top             =   30
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   926
         Caption         =   "【 TXX000001.TXT 】"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   375
            Left            =   2700
            Top             =   75
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   661
            Caption         =   "（ 99,999 / 99,999 ）"
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   360
         Left            =   30
         Top             =   630
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   635
         BackColor       =   14737632
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
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin CoReports.CrDraw CrDraw1 
      Left            =   120
      Top             =   1440
      _Version        =   131072
      _ExtentX        =   820
      _ExtentY        =   820
      _StockProps     =   0
   End
   Begin VB.Label PNL_Ok 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6945
      TabIndex        =   7
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
Attribute VB_Name = "CXJ0710"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  テキスト受注データ変換
'**       フォームID    :  CXJ0710
'**       処理概要      :
'**
'**             PARA - /M or /Manual  :手動処理
'**                    /D or /Disp    :画面表示あり(自動処理時のみ有効）
'**                    /R00000000     :再処理 00000000:履歴管理番号
'**
'**       作  成  日    :   2009/09/15
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
'
    Dim mSaveMsg        As String       ' ﾒｯｾｰｼﾞ

    Dim ZDate           As String
    Dim ZFSu            As Long
    Dim ZProc           As String
    
    Dim mKaraPrint      As Byte
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
'   #-------------------------#
'   #    テキスト 初期値セット   #
'   #-------------------------#
    Text_Proc = 0          ' 受注
    Text_ProcKB = 1        ' 手動
    Text_ProcDisp = 1      ' 画面表示あり
    Text_FunFlg = False
    
    If InStr(1, UCase(Command), "/AUTO") <> 0 Or InStr(1, UCase(Command), "/A") <> 0 Then
       Text_ProcKB = 0     ' 自動
       Text_ProcDisp = 0   ' 画面表示なし
    End If
    If Text_ProcKB = 0 Then
       If InStr(1, UCase(Command), "/DISP") <> 0 Or InStr(1, UCase(Command), "/D") <> 0 Then
          Text_ProcDisp = 1   ' 画面表示あり
       End If
    End If
    
    If InStr(1, UCase(Command), "/R") <> 0 Then
       Text_ProcKB = 1        ' 手動
       Text_ProcDisp = 1      ' 画面表示あり
       gStr = Replace(UCase(Command), "/R", "")
       Text_Rireki = CisFun.Val2(gStr)
    End If

    
    If Text_ProcKB = 0 Then
        If Not JTextYukou Then
           Call End_Proc
           Exit Sub
        End If
        CisDB.ErrDisp = False
    End If

    If Not JTextIniGet Then End
    
'   < シングルプロセスチェック(手動) >
    If Text_ProcKB <> 0 Then
       If Not SingleProcess_Check Then Call End_Proc
    End If
    
'   #-------------------#
'   # 対象ファイル取得  #
'   #-------------------#
    gLong = Text_GetFile_Count
    lb_FSu = Format(gLong, "#0")
'   #-------------------#
'   # 画面中央位置表示  #
'   #-------------------#
    If Text_ProcKB = 0 And Text_ProcDisp = 0 Then
       'Me.WindowState = 0
       Me.Left = Me.Width * -1
       Me.Top = Me.Height * -1
    Else
       Call CisFrm.Frm_Center(True, x1024y768)
    
       If gLogoDisp Then
          BG_Logo(0).Caption = gLogo
          CisFrm.BGLDisp BG_Logo, 10
       End If
    End If
    
    
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct
'< 空出荷指示書発行確認 >
    Call ItemsClearIRN
    mKaraPrint = 0
    IRN.bk = "Option": IRN.rno = 52
    If IRNRead("", 1) Then
       mKaraPrint = CisFun.Val2(Mid(IRN.oyian, 1, 1))
    End If
    
'< 前回処理日 >
    Call ZenProc_Get
    Hlb_ZDate = RTrim$(ZDate)
    If Trim(ZDate) <> "" Then
       Hlb_ZDate = Hlb_ZDate & ZProc
    End If
    Hlb_FSu = Format(ZFSu, "###")
    
' 自動処理
    If Text_ProcKB = 0 Then
       Timer1.Enabled = False
       Timer2.Enabled = False
       Call AutoProc
       Call PB_END_Click
    End If
' 再処理
    If Text_Rireki <> 0 Then
        
       Call ReturnPress
       Call PB_END_Click
    End If
    
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
    Call SingleProcess_Release
    
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
    B1_Area1.Visible = B1Mode

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
    gLong = Text_GetFile_Count
    lb_FSu = Format(gLong, "#0")
    
    If Text_Rireki = 0 Then
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
    End If
    ' 受注テキスト連携　環境設定確認
    If Not JTextIniCheck Then               ' 2009.09.15 upd : eKanbanIniCheck-->JTextIniCheck
       If Trim(Text_ErrMsg) <> "" Then
          With CisFun
               .MB_Lines = 7
               .MB_MSG(2) = "       テキスト変換環境設定エラー(受注)              "
               .MB_MSG(4) = "       " & Text_ErrMsg
               .MB_MSG(6) = "      テキスト変換環境設定を確認して下さい。         "
               .MB_Title = "テキスト変換環境設定ｴﾗｰ"
               .MB_Button = Error
               Call .MBOX
          End With
       End If
       GoTo ReturnPress_Ee
    End If
    
    If Text_Rireki = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "       処理を行います。         "
            .MB_MSG(4) = "       よろしいですか？         "
            .MB_Title = "テキスト受注ﾃﾞｰﾀ変換処理"
            .MB_Button = Yes_No
            If Not .MBOX Then GoTo ReturnPress_Ee
        End With
    End If
        
    Timer2.Enabled = False
    DoEvents
    
    On Error GoTo ReturnPress_Er
    
'------------------- 履歴処理開始
    If Not Start_TRIInsert(Text_Proc, Text_ProcKB, CisFun.Val2(lb_FSu)) Then
    
       GoTo ReturnPress_Er
    End If
    
    If Text_Rireki <> 0 Then
        gLong = Text_GetFile_Count
        lb_FSu = Format(gLong, "#0")
        
        If CisFun.Val2(gLong) = 0 Then
           With CisFun
                .MB_Lines = 3
                .MB_MSG(2) = "       処理対象ファイルがありません。              "
                .MB_Title = "受注テキスト処理"
                .MB_Button = OK
                Call .MBOX
           End With
           GoTo ReturnPress_Ee
        End If
    End If
    
    ' ログ出力処理開始
    Call Text_LogPut("", "S")
    
    Call DispChange("B1")
    '--------------------------------------
    '------------
    '------------       テキスト変換処理
    '------------
    '--------------------------------------
    If Not JTextProc Then GoTo ReturnPress_Er
'
ReturnPress_D:
'------------------------------------- 正常終了
    ' ログ出力処理終了
    Call Text_LogPut("", "E")
'
    gStr = "【  処　理  終  了  】"
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
    GoTo ReturnPress_Ed
ReturnPress_Er:
'------------------------------------- 異常終了
    ' ログ出力処理終了
    Call Text_LogPut("", "A")
    
    gStr = "【  処　理  異　常　終  了  】"
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
    GoTo ReturnPress_Ed
ReturnPress_Ee:
'------------------------------------- 再入力
    lb_msg = mSaveMsg
    DoEvents
    ProcHB = "H1"
    Call DispChange(ProcHB)
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/
'_/                 テキスト受注データ変換処理(自動)
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Private Sub AutoProc()
    
    Timer2.Enabled = False
    
    On Error GoTo AutoProc_Ee
    
    '------------------- 履歴処理開始
    If Not Start_TRIInsert(Text_Proc, Text_ProcKB, CisFun.Val2(lb_FSu)) Then
       GoTo AutoProc_Ee
    End If
    ' ログ出力処理開始
    Call Text_LogPut("", "S")
    
    ' シングルプロセスチェック
    If Not SingleProcess_Check Then
        gStr = "他のパソコンで起動中です       ( " & RTrim$(SYM.英数字1) & " )"
        Call Text_LogPut(gStr, "ERR")
        
        TRI.ERR = "04"      ' 重複起動ｴﾗｰ
        If Not Middle_TRIUpDate Then GoTo AutoProc_Ee
        
        ' ログ出力処理終了
        Call Text_LogPut("", "A")
        Call End_Proc
        Exit Sub
    End If
    
    ' テキスト環境設定確認
    If Not JTextIniCheck Then
        TRI.ERR = "05"      ' 環境設定ｴﾗｰ
        If Not Middle_TRIUpDate Then GoTo AutoProc_Ee
        GoTo AutoProc_Ee
    End If
    
    ' テキストデータ確認(該当データなし)
    If CisFun.Val2(lb_FSu) = 0 Then
          
        TRI.ERR = "01"
        If Not Middle_TRIUpDate Then GoTo AutoProc_Ee
        '+--------------------------- 今回処理日時を格納
        Call ZenProc_Put
         
        gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call Text_LogPut(gStr, , 1)
        gStr = "'"
        Call Text_LogPut(gStr, , 1)
        gStr = "＿／       【" & Text_PrefixNm(Text_Proc) & "】処理対象ファイルなし"
        Call Text_LogPut(gStr, , 1)
        gStr = "'"
        Call Text_LogPut(gStr, , 1)
        gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call Text_LogPut(gStr, , 1)
        gStr = "'"
        Call Text_LogPut(gStr, , 1)
    Else
        If Not JTextProc Then GoTo AutoProc_Ee
    End If
    ' ログ出力処理終了
    Call Text_LogPut("", "E")
    Exit Sub
AutoProc_Ee:
    ' ログ出力処理終了
    Call Text_LogPut("", "A")
End Sub
'*---------------------------------------------------------*
'*                  テキスト変換処理                       *
'*---------------------------------------------------------*
Private Function JTextProc() As Boolean
    On Error GoTo JTextProc_Er:
    
    JTextProc = False
'-------------------  テキストデータ移動先取得
    If Not JTextBackGet Then GoTo JTextProc_Er
    
    Call Text_LogProc(Text_Fun, "", "S")

'-------------------  テキストデータ移動処理
    If Text_Rireki = 0 Then
       If Text_ProcKB <> 0 Then lb_msg.Caption = "テキストデータ移動中": DoEvents
    
       If Not Text_TextMove(Me) Then GoTo JTextProc_Er
    End If
'-------------------  ﾃｷｽﾄｺﾝﾊﾞｰﾄ
    If Text_ProcKB <> 0 Then
        lb_msg.Caption = "テキスト変換中": DoEvents
        With CisFrm
            Set .MeterPic = Bar
            .MeterMax = CisFun.Val2(lb_FSu)
        End With
        Bar.Visible = True
        
        Set Text_Frm = CisFrm
    End If
    
    If Not Text_BULK_Insert Then GoTo JTextProc_Er
'------------------   テキスト受注更新処理
    If Text_ProcKB <> 0 Then lb_msg.Caption = "テキストデータ更新中": DoEvents
    
    If Not Text_UpdateProc Then GoTo JTextProc_Er
'------------------   テキスト受注累積処理
    If Text_ProcKB <> 0 Then lb_msg.Caption = "テキストデータ累積中": DoEvents
    
    If Not Jyucyu_Proc Then GoTo JTextProc_Er
'------------------   テキスト受注分納更新処理
' テキスト受注データに分納分はなし                   2008/07/18DEL
    If Text_ProcKB <> 0 Then lb_msg.Caption = "テキストデータ分納更新中": DoEvents
'    If Not Bun_Proc Then GoTo eKanbanProc_Er
'------------------   今回処理内容出力
    Call ZenProc_Put
'------------------- 履歴ﾃﾞｰﾀ削除
    If Text_DirMaxCnt = 1 Then
       If Text_ProcKB <> 0 Then lb_msg.Caption = "履歴削除中": DoEvents
       If Not Text_BackUpDel Then GoTo JTextProc_Er
    End If

''------------------- 出荷指示書発行
'    If Text_ProcKB <> 0 Then lb_msg.Caption = "出荷指示書発行": DoEvents
'    ' ログ出力
'    gStr = "【 出荷指示書発行 開始 】"
'    Call Text_LogPut(gStr, "", 1)
'
'    If mKaraPrint = 1 Then
'       gStr = "☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"
'       Call Text_LogPut(gStr, , 20)
'       gStr = "☆　　　　　　　　　　　　　　　　　　　　　☆"
'       Call Text_LogPut(gStr, , 20)
'       gStr = "☆☆☆☆☆☆ 空出荷指示書発行なし ☆☆☆☆☆☆"
'       Call Text_LogPut(gStr, , 20)
'       gStr = "☆　　　　　　　　　　　　　　　　　　　　　☆"
'       Call Text_LogPut(gStr, , 20)
'       gStr = "☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"
'       Call Text_LogPut(gStr, , 20)
'    End If
'
'    tri.ERR = "61"      '出荷指示書発行ｴﾗｰ
'    If Not Middle_triUpDate Then Exit Function
'
'    If Text_ProcKB = 0 And Text_ProcDisp = 0 Then
'        pref_DispForm = 1
'        Set pref_Form = CTP3110
'    End If
'    If Not CXJ0120LPrint(0, False) Then GoTo eKanbanProc_Er
'    ' ログ出力
'    gStr = "【 出荷指示書発行 終了 】"
'    Call Text_LogPut(gStr, "", 1)
'
'============================================= 2009/07/06
''------------------- 箱種一覧発行
'    If Text_ProcKB <> 0 Then lb_msg.Caption = "箱種一覧発行": DoEvents
'    ' ログ出力
'    gStr = "【 箱種一覧発行 開始 】"
'    Call Text_LogPut(gStr, "", 1)
'
'    tri.ERR = "62"      '箱種一覧発行ｴﾗｰ
'    If Not Middle_triUpDate Then Exit Function
'
'    If Text_ProcKB = 0 And Text_ProcDisp = 0 Then
'        pref_DispForm = 1
'        Set pref_Form = CTP3110
'    End If
'
'    If Not CXJ0125LPrint(0) Then GoTo eKanbanProc_Er
'    ' ログ出力
'    gStr = "【 箱種一覧発行 終了 】"
'    Call Text_LogPut(gStr, "", 1)
'============================================= 2009/07/06

'------------------- 変換リスト発行
    If Not Convert_ListProc Then GoTo JTextProc_Er
'-------------------  履歴出力(正常終了)
    TRI.ERR = ""
    If Not Middle_TRIUpDate Then GoTo JTextProc_Er  ' EHRUpdate 2009.09.15 memo
    
    JTextProc = True
    Exit Function
JTextProc_Er:

End Function
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*---------------                                                                             ---------------*
'*---------------       テキスト　受注累積処理                                                ---------------*
'*---------------                                                                             ---------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
Public Function Jyucyu_Proc() As Boolean
    On Error GoTo Jyucyu_Proc_Er:
    
    Jyucyu_Proc = False
    
    ' ログ出力
    gStr = "【 テキスト受注累積（出荷指示書NO採番) 開始 】"
    Call Text_LogPut(gStr, "", 1)
    
    TRI.ERR = "52"      'テキスト受注累積ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    With CisDB
        .SQL = "出荷指示書NO採番累積"
        .StoadoCount = 5
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = OutPut用
        .ParaNo = 3: .ParaIO = OutPut用
        .ParaNo = 4: .ParaIO = OutPut用
        .ParaNo = 5: .ParaIO = Return用
    '
        .ParaNo = 1: .ParaValue = 0
    
        If Not .DBStored Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo Jyucyu_Proc_Er
        End If
        .ParaNo = 5
        If .ParaValue <> 0 Then
            ReDim DispMsg(4)
            DispMsg(0) = "   テキスト受注累積エラー"
            DispMsg(2) = "   Err Code = " & .ParaValue
            GoTo Jyucyu_Proc_ErDisp
        End If
    End With
    
    ' ログ出力
    gStr = "【 テキスト受注累積（出荷指示書NO採番) 終了 】"
    Call Text_LogPut(gStr, "G", 1)
    
    Jyucyu_Proc = True
    Exit Function
Jyucyu_Proc_Er:
    ReDim DispMsg(4)
    DispMsg(0) = "   テキスト Update Err"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description

Jyucyu_Proc_ErDisp:
    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))

    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
    End If

    ' ログ出力
    gStr = "【 テキスト受注累積（出荷指示書NO採番) 異常終了 】"
    Call Text_LogPut(gStr, "", 1)
    
End Function
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*---------------                                                                             ---------------*
'*---------------     テキスト　受注分納更新               2008/07/18削除                     ---------------*
'*---------------                                                                             ---------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
Public Function Bun_Proc() As Boolean
    On Error GoTo Bun_Proc_Er:
    
    Bun_Proc = False
    
    ' ログ出力
    gStr = "【 テキスト受注分納更新 開始 】"
    Call Text_LogPut(gStr, "", 1)
    
    TRI.ERR = "53"      'テキストデータ納更新ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    With CisDB
        .SQL = "テキスト受注分納更新"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = OutPut用
        .ParaNo = 2: .ParaIO = OutPut用
        .ParaNo = 3: .ParaIO = Return用
    '
        If Not .DBStored Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo Bun_Proc_Er
        End If
        .ParaNo = 3
        If .ParaValue <> 0 Then
            ReDim DispMsg(4)
            DispMsg(0) = "   テキスト受注分納更新エラー"
            DispMsg(2) = "   Err Code = " & .ParaValue
            GoTo Bun_Proc_ErDisp
        End If
    
        ' ログ出力
        .ParaNo = 1
        gStr = " 入力対象件数 = " & Format(.ParaValue, "#,##0")
        Call Text_LogPut(gStr)
        .ParaNo = 2
        gStr = " 分納更新件数 = " & Format(.ParaValue, "#,##0")
        Call Text_LogPut(gStr)
    
    End With
    
    ' ログ出力
    gStr = "【 テキスト受注分納更新 終了 】"
    Call Text_LogPut(gStr, "G", 1)
    
    Bun_Proc = True
    Exit Function
Bun_Proc_Er:
    ReDim DispMsg(4)
    DispMsg(0) = "   テキスト Update Err"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description

Bun_Proc_ErDisp:
    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))

    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
    End If

    ' ログ出力
    gStr = "【 テキスト受注分納更新 異常終了 】"
    Call Text_LogPut(gStr, "", 1)
    
End Function
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*---------------                                                                             ---------------*
'*---------------     テキスト受注変換リスト処理                                              ---------------*
'*---------------                                                                             ---------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
Private Function Convert_ListProc() As Boolean
    Dim ErrStr  As String
    Dim KariStr As String
    
    
    On Error GoTo Convert_ListProc_Er:
    
    Convert_ListProc = False
    
    If Not Text_CvtList(Text_Proc) Then
       gStr = "☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"
       Call Text_LogPut(gStr, , 20)
       gStr = "☆　　　　　　　　　　　　　　　　　　　　　　　　　　     　 ☆"
       Call Text_LogPut(gStr, , 20)
       gStr = "☆☆☆☆☆☆　　　テキスト受注変換リスト発行なし　　☆☆☆☆☆☆"
       Call Text_LogPut(gStr, , 20)
       gStr = "☆　　　　　　　　　　　　　　　　　　　　　　　　　　     　 ☆"
       Call Text_LogPut(gStr, , 20)
       gStr = "☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"
       Call Text_LogPut(gStr, , 20)
       Convert_ListProc = True
       Exit Function
    End If
    If Text_ProcKB <> 0 Then lb_msg.Caption = "受注変換リスト発行": DoEvents
    
    ' ログ出力
    gStr = "【 テキスト受注変換リスト出力 開始 】"
    Call Text_LogPut(gStr, "", 1)
    
    TRI.ERR = "63"      '変換リスト発行ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function

' * 変換エラーリスト
    Text_ProcYMD = Format(Now(), "YYYYMMDD")
    
    Call Text_Cvt_List
    
    If THNRead(gSL_Select) Then
        
        lb_msg = "【 テキスト受注 】 変換リスト 印刷中"
        DoEvents
        With Text_Prt
             Set .iPrtForm = CXJ0710L 'CTP3110L
             .dpMSG = "テキスト受注変換リスト　印刷中"
             .iPreview = pv_Print
             If Text_ProcKB = 0 And Text_ProcDisp = 0 Then
                .dpMSGScreen = 画面なし
             Else
                .dpMSGScreen = モーダル
             End If
             
             .PrintStart
        End With
    End If
    Call THNClose
    
    ' ログ出力
    gStr = "【 テキスト受注変換リスト出力 終了 】"
    Call Text_LogPut(gStr, "", 1)
    
    Convert_ListProc = True
    Exit Function
Convert_ListProc_Er:
    ReDim DispMsg(4)
    DispMsg(0) = "   テキスト 変換リスト出力エラー  "
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description

    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))

    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
    End If

    ' ログ出力
    gStr = "【 テキスト受注変換リスト出力 異常終了 】"
    Call Text_LogPut(gStr, "", 1)


End Function


'*---------------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス チ ェ ッ ク            *
'*---------------------------------------------------------*
Private Function SingleProcess_Check()
    SingleProcess_Check = False
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "シングルプロセス"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "テキスト受注[愛知機械]"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)
    
    Call ItemsClearSYM
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称   = 'シングルプロセス'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値        = '" & App.EXEName & "'"
    Call SYMRead(gSL_Select, 1)
    
    If RTrim$(SYM.英数字1) <> "" And RTrim$(SYM.英数字1) <> CisFun.WSID Then
        With CisFun
            If Text_ProcKB <> 0 Then
                .MB_Lines = 3
                .MB_MSG(1) = "    他のパソコンで起動中です         "
                .MB_MSG(2) = "        ( " & RTrim$(SYM.英数字1) & " )"
                .MB_Button = Error
                .MBOX
            End If
        End With
        Exit Function
    End If

'   < プロセスのロック >
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = SubString( Host_Name(), 1, 20 ),"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

    SingleProcess_Check = True
End Function
'*-------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス 解 除           *
'*-------------------------------------------------*
Private Function SingleProcess_Release()
    SingleProcess_Release = False
    
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = '',"
    gSL_Select = gSL_Select & " 英数字2 = '',"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With
    
    SingleProcess_Release = True
End Function
'*-------------------------------------------------*
'*              前　回　処　理　日      ( 獲得 )
'*-------------------------------------------------*
Private Function ZenProc_Get()
    ZDate = ""
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "テキスト管理[愛知機械]"    ' 2009.09.15 upd : "新eかんばん管理"
        .区分桁数 = 2
        .値 = "JY"
        .制御区分 = "X"
        .値名称 = "テキスト受注[愛知機械]"      ' 2009.09.15 upd : "新eかんばん受注"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)
    
    Call ItemsClearSYM
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称   = 'テキスト管理[愛知機械]'"    ' 2009.09.15 upd : '新eかんばん管理'
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値        = 'JY'"
    Call SYMRead(gSL_Select, 1)
    
    ZDate = RTrim$(SYM.英数字2)
    If SYM.数字1 = 0 Then
       ZProc = "(自動)"
    Else
       ZProc = "(手動)"
    End If
    If SYM.数字3 <> 0 Then
       ZProc = "(再処理) "
    End If
    ZFSu = SYM.数字2
End Function
'*-------------------------------------------------*
'*              前　回　処　理　日      ( 書込 )
'*-------------------------------------------------*
Private Function ZenProc_Put()

    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = '',"
    gSL_Select = gSL_Select & " 英数字2 = '" & Format(Now(), "yyyy年mm月dd日 HH:MM") & "',"
    gSL_Select = gSL_Select & " 数字1 = " & Text_ProcKB & ","
    gSL_Select = gSL_Select & " 数字2 = " & CisFun.Val2(lb_FSu) & ","
    gSL_Select = gSL_Select & " 数字3 = " & Text_Rireki & ","
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'テキスト管理[愛知機械]'"  ' 2009.09.15 upd : '新eかんばん管理'
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = 'JY'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

End Function


