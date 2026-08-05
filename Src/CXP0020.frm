VERSION 5.00
Begin VB.Form CXP0020 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ﾊﾟｽﾜｰﾄﾞ設定"
   ClientHeight    =   1950
   ClientLeft      =   3705
   ClientTop       =   4260
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   7950
   Begin VB.CommandButton PB_ENT 
      Caption         =   "OK"
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
      Left            =   6810
      TabIndex        =   6
      Top             =   150
      Width           =   1035
   End
   Begin VB.CommandButton PB_Can 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   405
      Left            =   6810
      TabIndex        =   7
      Top             =   810
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "ﾊﾟｽﾜｰﾄﾞ大文字･小文字"
      Height          =   1695
      Left            =   3900
      TabIndex        =   9
      Top             =   90
      Width           =   2775
      Begin VB.OptionButton Op_PassType 
         Caption         =   "区別あり"
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   5
         Top             =   1110
         Width           =   1605
      End
      Begin VB.OptionButton Op_PassType 
         Caption         =   "区別なし"
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   450
         Value           =   -1  'True
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " ﾊﾟｽﾜｰﾄﾞ種類 "
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   90
      Width           =   3705
      Begin VB.OptionButton Op_Pass 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ﾊﾟｽﾜｰﾄﾞなし"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   3165
      End
      Begin VB.OptionButton Op_Pass 
         Caption         =   "ﾊﾟｽﾜｰﾄﾞのみ"
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   1
         Top             =   570
         Width           =   3165
      End
      Begin VB.OptionButton Op_Pass 
         Caption         =   "担当者 + ﾊﾟｽﾜｰﾄﾞ(個別)"
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   2
         Top             =   900
         Width           =   3165
      End
      Begin VB.OptionButton Op_Pass 
         Caption         =   "担当者 + ﾊﾟｽﾜｰﾄﾞ(共通)"
         Height          =   285
         Index           =   3
         Left            =   300
         TabIndex        =   3
         Top             =   1230
         Width           =   3165
      End
   End
End
Attribute VB_Name = "CXP0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  パスワード設定
'**       フォームID    :  CXP0020
'**       処理概要      :
'**
'**       作  成  日    :  2004/11/25  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    
    Dim CISFrm          As CISFormContorl

    Dim pStr            As String
    Dim pStr2           As String

    Dim mGotColor       As Long
    Dim mLostColor      As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   #  項 目 初 期 化  #
'   #------------------#
    Set CISFrm = New CISFormContorl
    Set CISFrm.ThisForm = Me
    Call CISFrm.InitFld
    
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CISFrm.Frm_Center(False, x1024y768)
    Me.Top = Me.Top - 2000
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    mGotColor = Op_Pass(0).BackColor
    mLostColor = Op_Pass(1).BackColor
    
    
    Call InitDisp
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    If Not FormAct Then Exit Function
    
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    
    Select Case Key_Code
        Case vbKeyEscape:   If PB_Can.Visible Then Call PB_Can_Click    '【終了】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                      '【入力】
                                Call PB_ENT_Click
                            End If
        Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Call PB_Can_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call PB_Can_Click

End Sub

'+---------------------+
'+    ｷｬﾝｾﾙ  キ ー     +
'+---------------------+
Private Sub PB_Can_Click()
    Me.Hide
End Sub
Private Sub Op_Pass_Click(Index As Integer)
    For Each gObj In Op_Pass
        gObj.BackColor = mLostColor
    Next gObj
    Op_Pass(Index).BackColor = mGotColor
End Sub
Private Sub Op_PassType_Click(Index As Integer)
    For Each gObj In Op_PassType
        gObj.BackColor = mLostColor
    Next gObj
    Op_PassType(Index).BackColor = mGotColor
End Sub


'+----------------------------+
'+      設定内容表示
'+----------------------------+
Private Sub InitDisp()
    ' パスワード種類設定
    With CisDB
        .SQL = "Select bk,rno,oyian From Irnaksic "
        .SQL = .SQL & " Where bk = 'SEn'"
        pStr = ""
        pStr2 = ""
        If .DBRead(1) Then
           gStr = RTrim(.Recordset(2))
           pStr = .PassWord(Mid(gStr, 1, 15))
           pStr2 = .PassWord(Mid(gStr, 16))
        End If
        Select Case RTrim(pStr)
            Case "PASS1": Op_Pass(1).Value = True
            Case "PASS2": Op_Pass(2).Value = True
            Case "PASS3": Op_Pass(3).Value = True
            Case Else: Op_Pass(0).Value = True
        End Select
        Select Case RTrim(pStr2)
            Case "0": Op_PassType(0).Value = True
            Case Else: Op_PassType(1).Value = True
        End Select
    End With
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    ' パスワード種類設定
    With CisDB
        If Op_Pass(0).Value Then gStr = "PASSNO"
        If Op_Pass(1).Value Then gStr = "PASS1"
        If Op_Pass(2).Value Then gStr = "PASS2"
        If Op_Pass(3).Value Then gStr = "PASS3"
        gStr = .PassWord(gStr)   ' パスワード種類
        If Op_PassType(0).Value Then pStr = "0"
        If Op_PassType(1).Value Then pStr = "1"
        pStr = .PassWord(pStr)   ' パスワードタイプ
        
        .SQL = "Select bk,rno,oyian From Irnaksic "
        .SQL = .SQL & " Where bk = 'SEn'"
        If Not .DBRead(1) Then
           
           .SQL = "Insert Into Irnaksic "
           .SQL = .SQL & " ( bk,rno,oyian,oyian2,etadpu) Values "
           .SQL = .SQL & " ( 'SEn',0,'" & gStr & pStr & "','',Getdate()) "
           If Not .DBExec Then GoTo PB_ENT_Er
        Else
           .SQL = "Update Irnaksic Set oyian = '" & gStr & pStr & "'"
           .SQL = .SQL & " Where bk = 'SEn'"
           If Not .DBExec Then GoTo PB_ENT_Er
        End If
    End With
    
'    Call MsgBox("設定終了           " & vbCrLf & "             ", vbOKOnly, "ﾊﾟｽﾜｰﾄﾞ設定")
    Call PB_Can_Click
    Exit Sub
PB_ENT_Er:
    MsgBox ("設定エラー")
End Sub

