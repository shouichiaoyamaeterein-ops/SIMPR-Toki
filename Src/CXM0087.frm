VERSION 5.00
Begin VB.Form CXM0087 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "パスワードの設定"
   ClientHeight    =   2310
   ClientLeft      =   5265
   ClientTop       =   3465
   ClientWidth     =   4950
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " 詳細設定 "
      Height          =   705
      Left            =   60
      TabIndex        =   9
      Top             =   2370
      Visible         =   0   'False
      Width           =   5955
      Begin VB.CheckBox Check1 
         Caption         =   "大文字小文字を区別する"
         Height          =   345
         Left            =   180
         TabIndex        =   10
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   2325
      End
   End
   Begin VB.CommandButton CmdCan 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   405
      Left            =   3630
      TabIndex        =   4
      Top             =   1020
      Width           =   1035
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3630
      TabIndex        =   3
      Top             =   360
      Width           =   1035
   End
   Begin VB.TextBox iH1_CheckPass 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1770
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "12345678901234567890"
      Top             =   1800
      Width           =   1665
   End
   Begin VB.TextBox iH1_NewPass 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1770
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "12345678901234567890"
      Top             =   1095
      Width           =   1665
   End
   Begin VB.TextBox iH1_OldPass 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1770
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "123456789012345"
      Top             =   390
      Width           =   1665
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "確認入力    ："
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1950
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "新しいパスワードの"
      Height          =   255
      Left            =   210
      TabIndex        =   7
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "新しいパスワード ："
      Height          =   255
      Left            =   210
      TabIndex        =   6
      Top             =   1140
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0C0C0&
      Caption         =   "古いパスワード  ："
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   450
      Width           =   1485
   End
End
Attribute VB_Name = "CXM0087"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  ﾊﾟｽﾜｰﾄﾞ設定
'**       フォームID    :  CXM0087
'**       処理概要      :
'**
'**       作  成  日    :  2003/11/17  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim wObj            As Object       ' ｵﾌﾞｼﾞｪｸﾄﾜｰｸｴﾘｱ
 
    Dim sDummySsap      As String
    Dim x               As String
    
    Dim mPChar          As Boolean

    Dim CISFrm          As CISFormContorl
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    
    c087_Rtn = False
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CISFrm = New CISFormContorl
    
    Set CISFrm.ThisForm = Me
    Call CISFrm.InitFld
    
    If c080_ProcKB <> "MNT" Then
       Me.Height = 2115
       Label1 = "パスワード："
       Label2 = "パスワード再入力："
    End If
    
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CISFrm.Frm_Center(False, x1024y768)
    
    x = CisDB.PassWord("")
    
    gSL_Select = "Select * from Irnaksic "
    gSL_Select = gSL_Select & " Where bk = '" & RTrim(c080_infoCD) & "'"
    gSL_Select = gSL_Select & " And   rno = 99"
    If IRNRead(gSL_Select, 1) Then
       x = Mid(IRN.oyian, 16)
    End If
    If c080_ProcKB = "MNT" Then
       If UCase(In_Pass) = "NAMCIS9611" Or UCase(In_Pass) = "CISADMIN" Then
          iH1_OldPass.PasswordChar = ""
          iH1_OldPass = CisDB.PassWord(x)
          x = Mid(IRN.oyian, 16)
       End If
    End If

End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub CmdOK_Click()
   Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub CmdCan_Click()
    Me.Hide
End Sub
''+----------------------------------------------------------+
''+      フォーム    《 QueryUnloadイベント 》                +
''+----------------------------------------------------------+
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    '
'    Call PB_CAN_Click
'End Sub
'+---------------------+
'+   iH1_CheckPass
'+---------------------+
Private Sub iH1_CheckPass_GotFocus()
    iH1_CheckPass.SelStart = 0
    iH1_CheckPass.SelLength = iH1_CheckPass.MaxLength
    iH1_CheckPass.IMEMode = iH1_CheckPass.IMEMode
End Sub
'+---------------------+
'+   iH1_NewPass
'+---------------------+
Private Sub iH1_NewPass_GotFocus()
    iH1_NewPass.SelStart = 0
    iH1_NewPass.SelLength = iH1_NewPass.MaxLength
    iH1_NewPass.IMEMode = iH1_NewPass.IMEMode
End Sub
'+---------------------+
'+   iH1_OldPass
'+---------------------+
Private Sub iH1_OldPass_GotFocus()
    iH1_OldPass.SelStart = 0
    iH1_OldPass.SelLength = iH1_OldPass.MaxLength
    iH1_OldPass.IMEMode = iH1_OldPass.IMEMode
End Sub
'+-----------------------------+
'+   ReturnPress(入力内容ﾁｪｯｸ)  +
'+-----------------------------+
Private Sub ReturnPress()
' 古いﾊﾟｽﾜｰﾄﾞ内容ﾁｪｯｸ
    sDummySsap = CisDB.PassWord(iH1_OldPass)
    
    If c080_ProcKB = "MNT" Then
        If sDummySsap <> x Then
            MsgBox "古いパスワードが間違っています", vbExclamation, "Pass"
            CmdOk.SetFocus
            iH1_OldPass.SetFocus
            Exit Sub
        End If
        If iH1_NewPass <> iH1_CheckPass Then
            MsgBox "パスワードエラー     ", vbExclamation, "パスワード"
            iH1_NewPass.SetFocus
            Exit Sub
        End If
    Else
        If iH1_OldPass <> iH1_NewPass Then
            MsgBox "パスワードエラー     ", vbExclamation, "パスワード"
            iH1_OldPass.SetFocus
            Exit Sub
        End If
    End If
    c080_Info = CisDB.PassWord(iH1_NewPass)
    
    With CisDB
        .SQL = "Update Irnaksic Set "
        .SQL = .SQL & " oyian = Substring(oyian,1,15) + '" & c080_Info & "'"
        .SQL = .SQL & " Where bk = '" & RTrim(c080_infoCD) & "'"
        .SQL = .SQL & " And   rno = 99"
        .DBExec
    End With
    
    c087_Rtn = True
    Call CmdCan_Click
    Exit Sub
ReturnPress_ER:
    CisFun.ErrorBox
End Sub

