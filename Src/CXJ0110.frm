VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Begin VB.Form CXJ0110 
   BackColor       =   &H00800000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "指示書№・納入番号採番"
   ClientHeight    =   3540
   ClientLeft      =   90
   ClientTop       =   1485
   ClientWidth     =   6825
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6825
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   525
      Left            =   540
      Top             =   390
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   926
      BackColor       =   8438015
      Caption         =   "※ 指示書№・納入番号採番 ※"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
   End
   Begin VB.TextBox Dummy 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Dummy"
      Top             =   0
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   420
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   2085
      Left            =   540
      ScaleHeight     =   2025
      ScaleWidth      =   5805
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1020
      Width           =   5865
      Begin Cis3D_v60.CIS3D CvtMsg 
         Height          =   585
         Left            =   450
         Top             =   510
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   1032
         BackColor       =   16776960
         Caption         =   "出荷指示書№採番 処理中"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   14.25
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
      Begin Cis3D_v60.CIS3D lb_Cnt 
         Height          =   405
         Left            =   2910
         Top             =   1470
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   714
         ForeColor       =   16711680
         Caption         =   "ZZ9"
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
         Height          =   405
         Left            =   1470
         Top             =   1470
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   714
         ForeColor       =   16711680
         Caption         =   "受注件数"
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
         cAlingnment     =   7
      End
   End
End
Attribute VB_Name = "CXJ0110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  出荷指示書№採番＆受注データ採番累積
'**       フォームID    :  CXJ0110
'**       処理概要      :
'**
'**       作  成  日    :  2004/03/22  By CIS
'**       変  更  日    :  2004/06/24  By CIS - 受領書・出荷時発行の対応
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld

    Dummy.Left = -1000
    Dummy.Enabled = True

    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = False   ' 時刻表示開始

'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(False)

End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    DoEvents
    FormAct = True

    Call ReturnPress
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)

    On Error GoTo ReturnPress_Err
'   << 対象データ確認 >>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select count(*) 件数"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " (     出荷指示書NO = 0"
    gSL_Select = gSL_Select & "    or 納入番号 = ''"
    gSL_Select = gSL_Select & "    or 累積区分 = 0"
    gSL_Select = gSL_Select & " )"
    If JNWRead(gSL_Select, 1) Then
        If JNW.件数 > 0 Then
            lb_Cnt.Caption = Format(JNW.件数, "#,##0")
            DoEvents
            With CisDB
            '   <><><> 出荷指示書NO採番＆受注テーブル累積 <><><>
'*--------------------------------------------------<< 2004/06/24 Update Start >>
'                .SQL = "出荷指示書NO採番累積"
'                .StoadoCount = 4
'                .ParaNo = 0: .ParaIO = Return用
'                .ParaNo = 1: .ParaIO = OutPut用
'                .ParaNo = 2: .ParaIO = OutPut用
'                .ParaNo = 3: .ParaIO = Return用
'                .ParaNo = 4: .ParaIO = Return用
'            '
'                If Not .DBStored Then
'                    MsgBox "ストアドエラー "
'                    GoTo ReturnPress_Err
'                Else
'                    .ParaNo = 4
'                    If .ParaValue <> 0 Then GoTo ReturnPress_Err
'                End If
                .SQL = "出荷指示書NO採番累積"
                .StoadoCount = 5
                .ParaNo = 0: .ParaIO = Return用
                .ParaNo = 1: .ParaIO = Input用
                .ParaNo = 2: .ParaIO = OutPut用
                .ParaNo = 3: .ParaIO = OutPut用
                .ParaNo = 4: .ParaIO = OutPut用
                .ParaNo = 5: .ParaIO = Return用
            '
                .ParaNo = 1: .ParaValue = gCHT_DelKbn53
            '
                If Not .DBStored Then
                    MsgBox "ストアドエラー "
                    GoTo ReturnPress_Err
                Else
                    .ParaNo = 5
                    If .ParaValue <> 0 Then GoTo ReturnPress_Err
                End If
'*--------------------------------------------------<< 2004/06/24 Update End >>
            End With
        End If
    End If
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
    Me.Hide
    Exit Sub
ReturnPress_Err:
    CisFun.ErrorBox
    End
End Sub
'*******************************
'*      メッセージ表示        *
'*******************************
Private Sub Timer1_Timer()
'''    Call ProcMsg(Timer1, CvtMsg, , 10, 100)
End Sub


