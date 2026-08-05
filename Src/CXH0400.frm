VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~1.OCX"
Begin VB.Form CXH0400 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "出力中"
   ClientHeight    =   1290
   ClientLeft      =   5100
   ClientTop       =   3990
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5640
   Begin Cis3D_v60.CIS3D Prt_Cnt 
      Height          =   330
      Left            =   1725
      Top             =   960
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   582
      BackColor       =   0
      ForeColor       =   65535
      Caption         =   "Z,ZZZ / Z,ZZZ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   8421504
      cFont3DColor2   =   -2147483643
      cFont3D         =   2
      cAlingnment     =   7
   End
   Begin CisMsg_v60.CISMSG Prt_Msg 
      Height          =   990
      Left            =   -15
      TabIndex        =   0
      Top             =   -45
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   1746
      BackColor       =   -2147483633
      ForeColor       =   16777215
      Caption         =   "【  出 力 中  】"
      BackStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   0
      cFont3DColor2   =   8421504
      cAlingnment     =   7
      gGC1_Red        =   200
      gGC2_Green      =   150
      gGC3_Bule       =   120
      gGradation      =   -1  'True
      MsgAction       =   3
      MsgActionTime   =   100
   End
End
Attribute VB_Name = "CXH0400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  自給チケットファイル出力
'**       フォームID    :  CXH0400
'**       処理概要      :
'**
'**       作  成  日    :  2005/02/22  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim mLoop           As Long
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

'   #----------------#
'   #    初期設定    #
'   #----------------#
    Prt_Msg = pref_Msg
    Prt_Cnt = ""
    FormAct = False

'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(False)
    DoEvents
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    Call CisFrm.MousePT(11)
    
    For mLoop = 1 To 200
        DoEvents
    Next mLoop
    
    Set pref_Cnt = Prt_Cnt

' 印刷処理
    Call JTicket_TextCreate

    Call CisFrm.MousePT(1)

' 処理終了
    Call ProcEnd
    
End Sub
'+----------------------+
'+      終了処理        +
'+----------------------+
Private Sub ProcEnd()
    Me.Hide
End Sub
