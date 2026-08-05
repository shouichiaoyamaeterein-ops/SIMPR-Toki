VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Begin VB.Form CNJ0315 
   BackColor       =   &H00C000C0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "増減・平準化更新"
   ClientHeight    =   3600
   ClientLeft      =   7875
   ClientTop       =   4245
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5460
   Begin VB.PictureBox H1_Area1 
      Height          =   2490
      Left            =   75
      ScaleHeight     =   2430
      ScaleWidth      =   5190
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   990
      Width           =   5250
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   510
         Index           =   0
         Left            =   15
         Top             =   390
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "かんばん回転枚数テーブル"
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
         cAlingnment     =   1
         cPositionX      =   90
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   510
         Index           =   1
         Left            =   15
         Top             =   900
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "かんばん増減テーブル"
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
         cAlingnment     =   1
         cPositionX      =   90
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   510
         Index           =   2
         Left            =   15
         Top             =   1410
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "かんばん平準化テーブル"
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
         cAlingnment     =   1
         cPositionX      =   90
      End
      Begin Cis3D_v60.CIS3D H1lb_Ins 
         Height          =   510
         Index           =   0
         Left            =   3000
         Top             =   390
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "999,999"
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
         cAlingnment     =   4
      End
      Begin Cis3D_v60.CIS3D H1lb_Ins 
         Height          =   510
         Index           =   1
         Left            =   3000
         Top             =   900
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "999,999"
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
         cAlingnment     =   4
      End
      Begin Cis3D_v60.CIS3D H1lb_Ins 
         Height          =   510
         Index           =   2
         Left            =   3000
         Top             =   1410
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "999,999"
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
         cAlingnment     =   4
      End
      Begin Cis3D_v60.CIS3D H1lb_Upd 
         Height          =   510
         Index           =   0
         Left            =   4095
         Top             =   390
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "999,999"
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
         cAlingnment     =   4
      End
      Begin Cis3D_v60.CIS3D H1lb_Upd 
         Height          =   510
         Index           =   1
         Left            =   4095
         Top             =   900
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "999,999"
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
         cAlingnment     =   4
      End
      Begin Cis3D_v60.CIS3D H1lb_Upd 
         Height          =   510
         Index           =   2
         Left            =   4095
         Top             =   1410
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "999,999"
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
         cAlingnment     =   4
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   510
         Index           =   9
         Left            =   15
         Top             =   1920
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "平準化更新エラーテーブル"
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
         cAlingnment     =   1
         cPositionX      =   90
      End
      Begin Cis3D_v60.CIS3D H1lb_Ins 
         Height          =   510
         Index           =   3
         Left            =   3000
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   12640511
         ForeColor       =   255
         Caption         =   "999,999"
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
         cAlingnment     =   4
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   510
         Index           =   11
         Left            =   4095
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         BackColor       =   12648447
         ForeColor       =   128
         Caption         =   ""
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
         cAlingnment     =   4
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            X1              =   105
            X2              =   960
            Y1              =   255
            Y2              =   255
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   390
         Index           =   12
         Left            =   15
         Top             =   0
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   688
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cPositionX      =   90
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   390
         Index           =   13
         Left            =   3000
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   688
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "追加"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   390
         Index           =   14
         Left            =   4095
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   688
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "更新"
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
         cAlingnment     =   7
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   885
      Left            =   45
      Top             =   45
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   1561
      BackColor       =   12648384
      ForeColor       =   0
      Caption         =   "増減 ・ 平準化情報更新"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cAlingnment     =   7
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   4500
         Top             =   615
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   12648384
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
         cBoderColor1    =   12648384
         cBoderColor2    =   12648384
         cFont3DColor1   =   8454143
         cFont3DColor2   =   16777215
         cFont3D         =   1
         cAlingnment     =   7
         cBoderStyle     =   2
      End
   End
End
Attribute VB_Name = "CNJ0315"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   増減・平準化更新
'**       フォームID    :   CNJ0315
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/15  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me

'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Call CisFrm.InitFld
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(False, x1024y768)
    
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    DoEvents
    If FormAct Then Exit Sub
    FormAct = Not FormAct

    Call ReturnPress
    
    Me.Hide
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)

    On Error GoTo ReturnPress_Err
    Call CisDB.DBTran(TransBegin)

'< 増減・平準化 >
    With CisDB
        .SQL = "増減平準データ作成"
        .StoadoCount = 10
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        
        For gInt = 4 To 10
            .ParaNo = gInt:    .ParaIO = OutPut用
        Next gInt
        
        .ParaNo = 1:    .ParaValue = g315_YM
        .ParaNo = 2:    .ParaValue = g315_ProcKB
        .ParaNo = 3:    .ParaValue = g315_FProc
    
        .DBStored
    
        .ParaNo = 0
        If .ParaValue >= 9 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    増減平準データ作成        "
            CisFun.MB_Button = Error
            CisFun.MBOX
            End
        End If
    
        .ParaNo = 4:    g315_InsCnt(0) = .ParaValue
        .ParaNo = 5:    g315_UpdCnt(0) = .ParaValue
        .ParaNo = 6:    g315_InsCnt(1) = .ParaValue
        .ParaNo = 7:    g315_UpdCnt(1) = .ParaValue
        .ParaNo = 8:    g315_InsCnt(2) = .ParaValue
        .ParaNo = 9:    g315_UpdCnt(2) = .ParaValue
    
        .ParaNo = 10:   g315_ErrCnt = .ParaValue
    End With
    
    For gInt = 0 To 2
        H1lb_Ins(gInt) = Format(g315_InsCnt(gInt), "#,###")
        H1lb_Upd(gInt) = Format(g315_UpdCnt(gInt), "#,###")
    Next gInt
    H1lb_Ins(3) = Format(g315_ErrCnt, "#,###")
    DoEvents
    
    Call CisDB.DBTran(TransCommit)
    On Error GoTo 0
    Call CisFrm.MousePT(1)
    
    Static sw_Cnt       As Long
    For sw_Cnt = 1 To 80000000: Next sw_Cnt
    
    Exit Sub

ReturnPress_Err:
    CisFun.ErrorBox "増減・平準化"
    Call CisFrm.MousePT(1)
    End
End Sub
