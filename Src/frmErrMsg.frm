VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Begin VB.Form frmErrMsg 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "通信ﾊﾟｯｹｰｼﾞｴﾗｰ情報更新"
   ClientHeight    =   1605
   ClientLeft      =   4575
   ClientTop       =   3495
   ClientWidth     =   5940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmErrMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5940
   Begin Cis3D_v60.CIS3D B1_TNSFMai 
      Height          =   555
      Left            =   270
      Top             =   150
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   979
      BackColor       =   0
      ForeColor       =   65535
      Caption         =   "通信ﾊﾟｯｹｰｼﾞｴﾗｰﾒｯｾｰｼﾞ更新中"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cFont3D         =   2
      cAlingnment     =   7
   End
   Begin VB.PictureBox B1_Gage 
      Height          =   675
      Left            =   270
      ScaleHeight     =   615
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   780
      Width           =   5355
   End
End
Attribute VB_Name = "frmErrMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  通信ﾊﾟｯｹｰｼﾞｴﾗｰ情報更新
'**       フォームID    :  frmErrMsg
'**
'**       作  成  日    :  2005/09/06  By CIS
'**       変  更  日    :
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
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(False, x1024y768)
    
End Sub
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    DoEvents
    
    Call TNSErrUpDate
    
    FormAct = Not FormAct
End Sub
'+------------------------------+
'+  通信ﾊﾟｯｹｰｰｼﾞｴﾗｰﾒｯｾｰｼﾞ更新    +
'+------------------------------+
Private Sub TNSErrUpDate()
    Static wFno     As Long
    Static wDate    As String
    Static wFCount  As Long
    
    Call CisFrm.MousePT(11)
    
    Call TETInsert
'
    On Error GoTo TNSErrUpDate_Er:
    
    wFno = FreeFile
    
    With CisFun
         .FE_FullPath = Ts_ErrFName
         wFCount = .FileLine
    End With
    
    With CisFrm
         Set .MeterPic = B1_Gage
         .MeterMax = wFCount
    End With
    
    gLong = 0
    
    Open Ts_ErrFName For Input As #wFno
    
    Do While Not EOF(wFno)          ' ファイルの終端までループを繰り返します。
        Line Input #1, wDate        ' 行を変数に読み込みます。
        gLong = gLong + 1
        If Trim(Mid(wDate, 1, 6)) <> "" Then
           With TET
                Call ItemsClearTET
                .ERRCODE = Mid(wDate, 1, 6)
                If Not TETRead("", 1) Then
                    .ERRLEBEL = Mid(wDate, 9, 1)
                    .ERRMSG1 = Replace(Trim(CisFun.Mid2(wDate, 14, 100)), "'", """")
                    .ERRMSG2 = Replace(Trim(CisFun.Mid2(wDate, 115, 98)), "'", """")
                    Call TETInsert
                Else
                    .ERRLEBEL = Mid(wDate, 9, 1)
                    .ERRMSG1 = Replace(Trim(CisFun.Mid2(wDate, 14, 100)), "'", """")
                    .ERRMSG2 = Replace(Trim(CisFun.Mid2(wDate, 115, 98)), "'", """")
                    Call TETUpdate
                End If
           End With
        End If
        CisFrm.MeterUpdate gLong
    Loop
    
    On Error GoTo 0
TNSErrUpDate_Er:
    Close wFno                            ' ファイルを閉じます。
    
    Call CisFrm.MousePT(1)
    
    Me.Hide
End Sub

