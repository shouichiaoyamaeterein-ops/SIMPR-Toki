VERSION 5.00
Begin VB.Form CXP0030 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "パスワードの変更"
   ClientHeight    =   2340
   ClientLeft      =   5475
   ClientTop       =   3480
   ClientWidth     =   6105
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " 詳細設定 "
      Height          =   705
      Left            =   60
      TabIndex        =   13
      Top             =   2370
      Visible         =   0   'False
      Width           =   5955
      Begin VB.CheckBox Check1 
         Caption         =   "大文字小文字を区別する"
         Height          =   345
         Left            =   180
         TabIndex        =   14
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   2325
      End
   End
   Begin VB.CommandButton CmdCan 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   405
      Left            =   5010
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
      Left            =   5010
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
      Left            =   3150
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
      Left            =   3150
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
      Left            =   3150
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "123456789012345"
      Top             =   390
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ﾊﾟｽﾜｰﾄﾞﾚﾍﾞﾙ"
      Height          =   2175
      Left            =   60
      TabIndex        =   8
      Top             =   90
      Width           =   1455
      Begin VB.OptionButton Op_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "管理責任者"
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   7
         Top             =   1650
         Width           =   1245
      End
      Begin VB.OptionButton Op_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "管理者"
         Height          =   345
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   1020
         Width           =   1245
      End
      Begin VB.OptionButton Op_Label 
         BackColor       =   &H00C0C0C0&
         Caption         =   "担当者"
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Tag             =   "KANRI"
         Top             =   390
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "確認入力    ："
      Height          =   255
      Left            =   1980
      TabIndex        =   12
      Top             =   1950
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "新しいパスワードの"
      Height          =   255
      Left            =   1590
      TabIndex        =   11
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "新しいパスワード ："
      Height          =   255
      Left            =   1590
      TabIndex        =   10
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "古いパスワード  ："
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   450
      Width           =   1425
   End
End
Attribute VB_Name = "CXP0030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  ﾊﾟｽﾜｰﾄﾞ変更
'**       フォームID    :  CXP0030
'**       処理概要      :
'**
'**       作  成  日    :  2004/11/25  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim wObj            As Object       ' ｵﾌﾞｼﾞｪｸﾄﾜｰｸｴﾘｱ
    
    Dim CISFrm          As CISFormContorl

    Dim sDummySsap      As String
    
    Dim mPChar          As Boolean

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    
'   #------------------#
'   #  項 目 初 期 化  #
'   #------------------#
    Set CISFrm = New CISFormContorl
    Set CISFrm.ThisForm = Me
    
    iH1_OldPass = ""
    iH1_NewPass = ""
    iH1_CheckPass = ""
    Op_Label(0).BackColor = &HFFFFC0
    
    mPChar = False
    If UCase(In_Pass) = "NAMCIS9611" Or UCase(In_Pass) = "CISADMIN" Then
        mPChar = True
        iH1_OldPass.PasswordChar = ""
    End If
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CISFrm.Frm_Center(False, x1024y768)
    
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '管理名'"
    gSL_Select = gSL_Select & "   And 値 <> ''"
    gSL_Select = gSL_Select & " Order By 値 "
    If SYMRead(gSL_Select) Then
       Do Until Not SYM_RDSTS
          gCnt1 = CisFun.Val2(SYM.値)
          If gCnt1 > 2 Then Exit Do
            
          Op_Label(gCnt1).Caption = Trim(SYM.値名称)
          Call SYMReadNext
       Loop
    End If
    Call SYMClose
    
    Erase P_Area
    
    For gCnt1 = 0 To 2
        gSL_Select = "Select bk,rno,oyian From Irnaksic "
        gSL_Select = gSL_Select & " Where bk = 'ssap'"
        gSL_Select = gSL_Select & "   And rno = " & gCnt1
        If IRNRead(gSL_Select, 1) Then
           P_Area(gCnt1) = RTrim(IRN.oyian)
           P_Area(gCnt1) = Trim(CisDB.PassWord(P_Area(gCnt1)))
        End If
    Next gCnt1

    Op_Label(0).Value = True

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
Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Call CmdCan_Click
End Sub
'+---------------------+
'+   Op_Label
'+---------------------+
Private Sub Op_Label_Click(Index As Integer)
    If mPChar Then
        iH1_OldPass = P_Area(Index)
    End If
    iH1_OldPass.SetFocus
    For Each wObj In Op_Label
        If Index = wObj.Index Then
            wObj.BackColor = &HFFFFC0
        Else
            wObj.BackColor = &HC0C0C0
        End If
    Next wObj
End Sub
'+---------------------+
'+   iH1_CheckPass
'+---------------------+
Private Sub iH1_CheckPass_GotFocus()
 '   iH1_CheckPass.BackColor = gIGotSel
    iH1_CheckPass.SelStart = 0
    iH1_CheckPass.SelLength = iH1_CheckPass.MaxLength
    iH1_CheckPass.IMEMode = iH1_CheckPass.IMEMode
End Sub
Private Sub iH1_CheckPass_LostFocus()
 '   iH1_CheckPass.BackColor = gILostSel
End Sub
'+---------------------+
'+   iH1_NewPass
'+---------------------+
Private Sub iH1_NewPass_GotFocus()
'    iH1_NewPass.BackColor = gIGotSel
    iH1_NewPass.SelStart = 0
    iH1_NewPass.SelLength = iH1_NewPass.MaxLength
    iH1_NewPass.IMEMode = iH1_NewPass.IMEMode
End Sub
Private Sub iH1_NewPass_LostFocus()
'    iH1_NewPass.BackColor = gILostSel
End Sub
'+---------------------+
'+   iH1_OldPass
'+---------------------+
Private Sub iH1_OldPass_GotFocus()
'    iH1_OldPass.BackColor = gIGotSel
    iH1_OldPass.SelStart = 0
    iH1_OldPass.SelLength = iH1_OldPass.MaxLength
    iH1_OldPass.IMEMode = iH1_OldPass.IMEMode
End Sub
Private Sub iH1_OldPass_LostFocus()
'    iH1_OldPass.BackColor = gILostSel
End Sub
'+-----------------------------+
'+   ReturnPress(入力内容ﾁｪｯｸ)  +
'+-----------------------------+
Private Sub ReturnPress()
    Static wPass    As Boolean
    Static wID      As Integer
    Dim sData       As String
' 古いﾊﾟｽﾜｰﾄﾞ内容ﾁｪｯｸ
    sDummySsap = iH1_OldPass
'    If sDummySsap = "" Then sDummySsap = sSAPCPS
    wID = 0
    For gInt = 0 To 2
        If Op_Label(gInt).Value Then
            gStr = P_Area(gInt)
            Exit For
        End If
        wID = wID + 1
    Next gInt
'''
    wPass = False
    For gInt = wID To 2
        If sDummySsap = P_Area(gInt) Then
           wPass = True
           Exit For
        End If
    Next gInt

    If Not wPass Then
        MsgBox "古いパスワードが間違っています。      ", vbExclamation, "Pass"
        CmdOk.SetFocus
        iH1_OldPass.SetFocus
        Exit Sub
    End If
    If iH1_NewPass <> iH1_CheckPass Then
        MsgBox "新しいﾊﾟｽﾜｰﾄﾞと確認入力内容が違います。      ", vbExclamation, "パスワード"
        iH1_NewPass.SetFocus
        Exit Sub
    End If
'''
    On Error GoTo ReturnPress_ER:

    For gInt = 0 To 2
        If Op_Label(gInt).Value Then
            If Trim(iH1_NewPass) <> "" Then
               gStr = iH1_NewPass
            End If
            If Trim(iH1_NewPass) = "" Then
               gStr = Space(15)
            End If
            gStr = gStr & Space(15)
            sData = CisDB.PassWord(Mid(gStr, 1, 15))
            Call ItemsClearIRN
            With IRN
                .bk = "ssap"
                .rno = gInt
            End With
            If IRNRead("", 1) Then
               IRN.oyian = sData
               Call IRNUpdate
            Else
               IRN.oyian = sData
               Call IRNInsert
            End If
        End If
    Next gInt

    Call CmdCan_Click
    Exit Sub
ReturnPress_ER:
    MsgBox Err & ":" & Error
End Sub
