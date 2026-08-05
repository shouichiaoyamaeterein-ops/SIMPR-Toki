VERSION 5.00
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~1.OCX"
Object = "{A37A80A5-3F34-11D2-912B-00C0DF49BE5E}#2.0#0"; "CrDraw.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form CXM0000L 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "印刷中"
   ClientHeight    =   945
   ClientLeft      =   7020
   ClientTop       =   4755
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   5595
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   570
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
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
      gGC1_Red        =   100
      gGC2_Green      =   250
      gGC3_Bule       =   250
      gGradation      =   -1  'True
   End
   Begin CoReports.CrDraw CrDraw1 
      Left            =   4650
      Top             =   480
      _Version        =   131072
      _ExtentX        =   820
      _ExtentY        =   820
      _StockProps     =   0
   End
End
Attribute VB_Name = "CXM0000L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  環境設定 発行
'**       フォームID    :  CXM0000L
'**       処理概要      :
'**
'**       作  成  日    :  2006/04/03  By CIS
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

    Set CXM0000_MailSes = MAPISession1
    Set CXM0000_MailMsg = MAPIMessages1
    
'   #----------------#
'   #    初期設定    #
'   #----------------#
    Prt_Msg = pref_Msg
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
    
' 印刷処理
    Call CoPrintM0000

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
