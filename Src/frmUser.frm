VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ﾕｰｻﾞｰ情報"
   ClientHeight    =   1905
   ClientLeft      =   5190
   ClientTop       =   3810
   ClientWidth     =   4170
   ControlBox      =   0   'False
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCompany 
      Height          =   345
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   0
      Top             =   300
      Width           =   675
   End
   Begin VB.TextBox txtCompName 
      Height          =   345
      IMEMode         =   4  '全角ひらがな
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   1
      Top             =   810
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   810
      TabIndex        =   2
      Top             =   1350
      Width           =   945
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   345
      Left            =   2250
      TabIndex        =   3
      Top             =   1350
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "ﾕｰｻﾞｰｺｰﾄﾞ　："
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "ﾕｰｻﾞｰ名     ："
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   870
      Width           =   1485
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  ﾕｰｻﾞｰ情報設定
'**       フォームID    :  frmUser
'**       処理概要      :
'**
'**       作  成  日    :  2001/10/10  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim mRtn            As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Call CisFrm.Frm_Center(False)
    txtCompany = Trim(gCompany)
    txtCompName = Trim(gCompName)
End Sub
'+----------------------------+
'+  　　　確　　定           　+
'+----------------------------+
Private Sub CmdOK_Click()
    gCompany = txtCompany
    If Trim(txtCompany) = "" Then
       MsgBox "ユーザーコードを入力して下さい。", vbExclamation, "必須入力ｴﾗｰ"
       txtCompany.SetFocus
       Exit Sub
    End If
    gCompName = txtCompName
    If Trim(gCompName) = "" Then
       mRtn = MsgBox("ユーザー名が入力されていません。       " & vbCrLf & vbCrLf & _
              "        よろしいですか？                      ", vbExclamation + vbYesNo, "入力確認")
       If mRtn <> vbYes Then
          txtCompName.SetFocus
          Exit Sub
       End If
    End If
' ﾕｰｻﾞｰ情報更新
    If Not CisDB.DBConnect(False) Then
    End If
'' ﾕｰｻﾞｰｺｰﾄﾞ & ﾛｺﾞ表示
    Call ItemsClearIRN
    IRN.bk = "Company": IRN.rno = 0
    If IRNRead("", 1) Then
       IRN.oyian = Trim(gCompany)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(gCompany)
       IRN.oyian2 = "SIMPR-A"
       Call IRNInsert
    End If
' ﾕｰｻﾞｰ名
    Call ItemsClearIRN
    IRN.bk = "Company": IRN.rno = 1
    If IRNRead("", 1) Then
       IRN.oyian = Trim(gCompName)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(gCompName)
       Call IRNInsert
    End If

    Me.Hide
End Sub
'+----------------------------+
'+  　　　取り消し           　+
'+----------------------------+
Private Sub cmdCan_Click()
    gInt = MsgBox("終了します、よろしいですか？？", vbOKCancel + vbExclamation, "ﾕｰｻﾞｰ情報")
    If gInt = vbOK Then
       Call CisDB.DBDISConnect
       Call CisFrm.UnLoadDisp
       End
    End If
End Sub
'+----------------------------+
'+ 　ユーザー名称自動表示       +
'+----------------------------+
Private Sub txtCompany_LostFocus()
    If Trim(txtCompName) <> "" Then Exit Sub
    If txtCompany = "9611" Then
        txtCompName = "㈱中京情報システム"
    End If
End Sub
