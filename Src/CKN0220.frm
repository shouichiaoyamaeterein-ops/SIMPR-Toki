VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Begin VB.Form CKN0220 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注かんばん 発行"
   ClientHeight    =   9600
   ClientLeft      =   255
   ClientTop       =   1170
   ClientWidth     =   14310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9600
   ScaleWidth      =   14310
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   555
      Left            =   5760
      Top             =   6705
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      Caption         =   "終  了"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   15.75
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
      Height          =   3030
      Left            =   4260
      ScaleHeight     =   3091.228
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6765
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3465
      Width           =   6825
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   2850
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6675
         _ExtentX        =   11774
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
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   780
            Left            =   330
            Top             =   915
            Width           =   6000
            _ExtentX        =   10583
            _ExtentY        =   1376
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
            cBoderColor1    =   64
            cBoderColor2    =   0
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderStyle     =   2
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   240
               Left            =   2940
               Top             =   270
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   423
               Caption         =   "～"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9
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
            Begin CisText_V60.CisText iH1_SNo 
               Height          =   360
               Left            =   2190
               TabIndex        =   6
               Top             =   210
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   635
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
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   3
            End
            Begin Cis3D_v60.CIS3D UC_3D2 
               Height          =   390
               Index           =   1
               Left            =   1035
               Top             =   195
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   688
               ForeColor       =   16711680
               Caption         =   "印刷頁"
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
            Begin CisText_V60.CisText iH1_ENo 
               Height          =   360
               Left            =   3300
               TabIndex        =   7
               Top             =   195
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   635
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
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   375
            Left            =   3420
            Top             =   1875
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   661
            Caption         =   "未発行枚数"
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin Cis3D_v60.CIS3D lb_HSu 
               Height          =   330
               Left            =   1350
               Top             =   15
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   582
               BackColor       =   16777152
               Caption         =   "#,### / #,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D28 
            Height          =   600
            Index           =   1
            Left            =   60
            Top             =   60
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   1058
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
               Height          =   510
               Left            =   45
               ScaleHeight     =   450
               ScaleWidth      =   6375
               TabIndex        =   2
               TabStop         =   0   'False
               Top             =   30
               Width           =   6435
               Begin Cis3D_v60.CIS3D lb_msg 
                  Height          =   420
                  Left            =   60
                  Top             =   60
                  Width           =   6255
                  _ExtentX        =   11033
                  _ExtentY        =   741
                  BackColor       =   16761087
                  Caption         =   "発注かんばんの発行を行います"
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   375
            Left            =   3420
            Top             =   2250
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   661
            Caption         =   "発行済枚数"
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin Cis3D_v60.CIS3D lb_SSu 
               Height          =   330
               Left            =   1350
               Top             =   15
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "#,### / #,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
         End
      End
   End
   Begin Cis3D_v60.CIS3D Head 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 発注かんばん 発行 】"
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
         Left            =   14130
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
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
         Height          =   300
         Left            =   11625
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   0
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
      TabIndex        =   1
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   555
      Left            =   7995
      Top             =   6705
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      Caption         =   "開  始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   15.75
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
   Begin Cis3D_v60.CIS3D PB_Kanban 
      Height          =   540
      Left            =   12600
      Top             =   435
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   953
      ForeColor       =   16711935
      Caption         =   "発行品番参照(F5)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   12
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
   Begin Cis3D_v60.CIS3D PB_OK 
      Height          =   555
      Left            =   6900
      Top             =   6690
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      Caption         =   "確　認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   15.75
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
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6930
      TabIndex        =   8
      Top             =   7260
      Width           =   1515
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8025
      TabIndex        =   5
      Top             =   7320
      Width           =   1515
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5775
      TabIndex        =   4
      Top             =   7320
      Width           =   1515
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
      TabIndex        =   3
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CKN0220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばん 発行
'**       フォームID    :  CKN0220
'**       処理概要      :
'**
'**       作  成  日    :  2003/12/24
'**       変  更  日    :
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
    
    Dim mSelColor       As Long
    Dim mUnSelColor     As Long
    Dim mSavePrint      As String
    
    Dim mSaveMsg        As String
    
    Dim mPrtKensu       As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If kn_Call = "" Then
    '   #------------------#
    '   # 二 重 起 動 防 止 #
    '   #------------------#
        If Not Dupli_Start Then End
    End If
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

    If kn_Call = "" Then
        Select Case App.Title
            Case "PKN0250": kn_PrtKB = 1
            Case Else
                 kn_PrtKB = 0
        End Select
    End If
    Select Case kn_PrtKB
        Case 1
            lb_msg = "工程かんばんの発行を行います"
            Head = "【 工程かんばん 発行 】"
        Case Else
            lb_msg = "発注かんばんの発行を行います"
            Head = "【 発注かんばん 発行 】"
    End Select
    
    With CisFrm
        Set .MSGObject = lb_msg
        Set .MSGTimer = Timer2
        .MSGStopCount = 10
        .MSGType = Left_To_Right
        .MSGText = lb_msg
        .MSGInterval = 100
        .MSGRelate
    End With
    mSaveMsg = lb_msg

    Dummy.Left = -1000
    Dummy.Enabled = False
   
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
    If kn_Call = "" Then
       Me.Show
    Else
       PB_END.Caption = "戻　る"
    End If
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    Call MaisuDisp

    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始
    
    FormAct = Not FormAct
    
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
        Case vbKeyF5: If PB_Kanban.Visible Then Call PB_Kanban_Click
        Case vbKeyF12
            If PB_END.Visible Then Call PB_END_Click        '【終了】
            If PB_OK.Visible Then Call PB_END_Click         '【終了】
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
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    If kn_Call = "" Then
       Cancel = 1
    End If
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'ｸﾞﾗﾃﾞｰｼｮﾝ表示
    With CisFrm
        .Gradation_Direction = Landscape
        If kn_PrtKB = 0 Then
           .Gradation_Shading = Lighten
           .Gradation 80, 100, 120, 2, 2, 3
        Else
           .Gradation_Shading = Deepen
           .Gradation 80, 100, 100, 2, 2, 3
        End If
    End With
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If kn_Call = "" Then
        Call CisDB.DBDISConnect
        Call CisFrm.UnLoadDisp
        
        End
    Else
        Unload CKN0220
    End If
End Sub
'+----------------------------+
'+    発行品番参照キ ー(F4)    +
'+----------------------------+
Private Sub PB_Kanban_Click()
    CKN0225.Show vbModal
    Unload CKN0225
    Set CKN0225 = Nothing

    Call MaisuDisp
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
       Case "H1"
          H1Mode = True
       Case "B1"
          B1Mode = True
       Case "B2"
          B2Mode = True
    End Select
    H1_Area1.Enabled = H1Mode
    PB_ENT.Visible = H1Mode
    PNL_ENT.Visible = H1Mode
    PB_END.Visible = H1Mode
    PNL_END.Visible = H1Mode
    Timer2.Enabled = H1Mode Or B2Mode
    PB_OK.Visible = B2Mode
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
    
    If Not Head1Chk Then GoTo ReturnPress_Ee
        
    Call DispChange("B1")
    
    If Not PrintStart Then GoTo ReturnPress_Ee

    gStr = "【  出  力  終  了  】"
    lb_msg.Caption = gStr
    DoEvents
    
    Call CisFrm.MousePT(1)
    With CisFrm
        .MSGType = FSize_Changes
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
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Dim wSl_Where   As String
    
    Head1Chk = False
    
    If CisFun.Val2(iH1_SNo) > mPrtKensu Or CisFun.Val2(iH1_ENo) > mPrtKensu Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　印刷頁は発行枚数以下にて入力して下さい    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_ENo.SetFocus
        Exit Function
    End If
    If CisFun.Val2(iH1_SNo) <> 0 And CisFun.Val2(iH1_ENo) <> 0 Then
        If CisFun.Val2(iH1_SNo) > CisFun.Val2(iH1_ENo) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　印刷頁を正しく入力して下さい（開始＞終了）       "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_SNo.SetFocus
            Exit Function
        End If
    End If

    lb_msg.Caption = "【 かんばん発行データ 抽出中 】"
    DoEvents
    
    gSL_Select = "Select * From かんばん発行テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & kn_PrtKB
    gSL_Select = gSL_Select & " Order By 発行種類,取引先,受入,品番,背番号,収容区分,発行NO"
    If Not KHTRead(gSL_Select, , 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　かんばん発行データがありません       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SNo.SetFocus
        Exit Function
    End If

    Kn_PCnt = CisDB.RecordCount

    Head1Chk = True
End Function
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
'+---------------------------------------------+
'+      かんばん枚数を求める
'+---------------------------------------------+
Private Sub MaisuDisp()
    ' 発行枚数を求める
    gSL_Select = "Select Count(*) 件数 From かんばん発行テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & kn_PrtKB
    gSL_Select = gSL_Select & "   And 発行済 = 0"
    If KHTRead(gSL_Select, 1) Then
        mPrtKensu = KHT.件数
        lb_HSu = "       " & CisFun.RSetFld(mPrtKensu, 5, "#,###")
    End If
    gSL_Select = "Select Count(*) 件数 From かんばん発行テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & kn_PrtKB
    gSL_Select = gSL_Select & "   And 発行済 = 1"
    If KHTRead(gSL_Select, 1) Then
        mPrtKensu = mPrtKensu + KHT.件数
        lb_SSu = "       " & CisFun.RSetFld(KHT.件数, 5, "#,###")
    End If
End Sub
'+---------------------------------------------+
'+      印刷処理                               +
'+---------------------------------------------+
Private Function PrintStart() As Boolean
    PrintStart = False
    
    lb_msg.Caption = "【 かんばん発行データ 印刷中 】"
    DoEvents
    Kn_Page = True
    Kn_PageS = CisFun.Val2(iH1_SNo)
    Kn_PageE = CisFun.Val2(iH1_ENo)
    
    CKN0240.Show vbModal
    Unload CKN0240
    Set CKN0240 = Nothing
    
    If Not Kn_Rtn Then Exit Function
    
    PrintStart = True
End Function
