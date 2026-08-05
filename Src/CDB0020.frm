VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Begin VB.Form CDB0020 
   BorderStyle     =   0  'なし
   ClientHeight    =   555
   ClientLeft      =   2505
   ClientTop       =   5175
   ClientWidth     =   10170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin Cis3D_v60.CIS3D ExecMsg 
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   979
      Caption         =   "データベース処理中"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   -2147483643
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
   End
End
Attribute VB_Name = "CDB0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  データベース処理中画面
'**       フォームID    :  CDB0020
'**       処理概要      :
'**
'**       作  成  日    :  2004/03/11  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CISFrm          As CISFormContorl

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CISFrm = New CISFormContorl
    
    Set CISFrm.ThisForm = Me
    
    ExecMsg = db_ExecMsg
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CISFrm.Frm_Center(False, x1024y768)
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
End Sub

