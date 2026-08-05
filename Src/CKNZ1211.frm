VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~2.OCX"
Object = "{A37A80A5-3F34-11D2-912B-00C0DF49BE5E}#2.0#0"; "CrDraw.ocx"
Begin VB.Form CKNZ1211 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "印刷中"
   ClientHeight    =   1770
   ClientLeft      =   6225
   ClientTop       =   4560
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5625
   Begin VB.CommandButton PB_Can 
      Caption         =   "印刷中止"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1980
      TabIndex        =   1
      Top             =   1260
      Width           =   1635
   End
   Begin Cis3D_v60.CIS3D Prt_Cnt 
      Height          =   330
      Left            =   1725
      Top             =   810
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   582
      BackColor       =   16777152
      ForeColor       =   0
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
      Height          =   1800
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   3175
      BackColor       =   -2147483633
      ForeColor       =   16777215
      Caption         =   "【  印 刷 中  】"
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
      cPositionY      =   -400
      gGC1_Red        =   150
      gGC2_Green      =   100
      gGC3_Bule       =   200
      gGradation      =   -1  'True
      Begin CoReports.CrDraw CrDraw1 
         Left            =   150
         Top             =   120
         _Version        =   131072
         _ExtentX        =   820
         _ExtentY        =   820
         _StockProps     =   0
      End
   End
End
Attribute VB_Name = "CKNZ1211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  発注かんばん 発行
'**       フォームID    :  CKNZ1211L
'**       処理概要      :
'**
'**       作  成  日    :  2006/08/09  By CIS CoReports8.5
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
    
    pref_Cansel_Flg = False
    
    For mLoop = 1 To 200
        DoEvents
    Next mLoop
    
    Set pref_Cnt = Prt_Cnt
    
    
' 印刷処理
    pref_Rtn = False
    If CoPrintN1230 Then
       pref_Rtn = True
    End If
    
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
'+----------------------------+
'+    印刷取消キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    pref_Cansel_Flg = True
End Sub
