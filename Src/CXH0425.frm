VERSION 5.00
Begin VB.Form CXH0425 
   Caption         =   "Form1"
   ClientHeight    =   1155
   ClientLeft      =   4620
   ClientTop       =   4560
   ClientWidth     =   7020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "発注入荷状況問合せ　起動中"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   435
      TabIndex        =   0
      Top             =   345
      Width           =   6165
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "発注入荷状況問合せ　起動中"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   465
      TabIndex        =   1
      Top             =   405
      Width           =   6165
   End
End
Attribute VB_Name = "CXH0425"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************'
'**
'**       フォーム名    :   発注計画状況照会
'**       フォームID    :   CXH0425
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/18  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Me.Caption = ""
    
    Call CisFrm.Frm_Center(False)
    
    Me.Show
    
    DoEvents
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'ｸﾞﾗﾃﾞｰｼｮﾝ表示
    With CisFrm
        .Gradation_Direction = Landscape
        .Gradation_Shading = Lighten
        .Gradation 80, 100, 180, 2, 2, 3
    End With
End Sub

