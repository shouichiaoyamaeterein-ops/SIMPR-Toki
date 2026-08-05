VERSION 5.00
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Begin VB.Form CCT0041 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "TEXTﾌｧｲﾙ追加"
   ClientHeight    =   3195
   ClientLeft      =   4455
   ClientTop       =   3015
   ClientWidth     =   7080
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7080
   Begin VB.CommandButton PB_CAN 
      Caption         =   "戻る"
      Height          =   345
      Left            =   4290
      TabIndex        =   5
      Top             =   2715
      Width           =   960
   End
   Begin VB.CommandButton PB_ENT 
      Caption         =   "次へ"
      Height          =   330
      Left            =   2010
      TabIndex        =   4
      Top             =   2745
      Width           =   960
   End
   Begin VB.PictureBox Picture2 
      Height          =   1560
      Left            =   165
      ScaleHeight     =   1500
      ScaleWidth      =   6750
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1050
      Width           =   6810
      Begin VB.OptionButton Op_TnsType 
         Caption         =   "手配データ(TB：020,TY:112)"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   0
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1110
         Width           =   6735
      End
      Begin VB.OptionButton Op_TnsType 
         Caption         =   "そ　　の　　他"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   8
         Left            =   360
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   9
         Tag             =   "0"
         Top             =   1560
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.OptionButton Op_TnsType 
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   3555
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   8
         Tag             =   "0"
         Top             =   1560
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.OptionButton Op_TnsType 
         Caption         =   "財 務 デ ー タ"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   0
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   3
         Tag             =   "1"
         Top             =   720
         Width           =   6735
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   345
         Left            =   -15
         Top             =   -15
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   609
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "TEXT変換種類"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   2
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin VB.OptionButton Op_TnsType 
         Caption         =   "内 示 デ ー タ"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   0
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   2
         Tag             =   "1"
         Top             =   330
         Width           =   6735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   870
      Left            =   165
      ScaleHeight     =   810
      ScaleWidth      =   6750
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   180
      Width           =   6810
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   810
         Left            =   1515
         Top             =   0
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "業　務　名　（TEXTファイル名称)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_GyoNm 
            Height          =   345
            Left            =   120
            TabIndex        =   1
            Top             =   330
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
            MaxLength       =   40
            IMEMode         =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   810
         Left            =   0
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "業務コード"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Gyomu 
            Height          =   345
            Left            =   120
            TabIndex        =   0
            Top             =   345
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXX"
            MaxLength       =   9
         End
      End
   End
End
Attribute VB_Name = "CCT0041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   TEXTﾌｧｲﾙ追加
'**       フォームID    :   CCT0041
'**       処理概要      :
'**
'**       作  成  日    :   2004/11/08  By CIS
'**       変  更  日    :   2007/03/05  By CIS 手配データ追加
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ

    Dim CisFrm          As CISFormContorl
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    Op_TnsType(0).Value = True  ' 内示初期選択
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(False, x1024y768)
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormAct = Not FormAct
'
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                  Call PB_ENT_Click
                               End If
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function

'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Me.Hide
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    If Trim(iH1_Gyomu) = "" Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "   業務コードを入力して下さい。     "
            .MB_Title = ""
            .MB_Button = Error
            Call .MBOX
       End With
       iH1_Gyomu.SetFocus
       Exit Sub
    End If
    TKT.FILECD = iH1_Gyomu
    If TKTRead("", 1) Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "   業務コードはすでに登録されています。     "
            .MB_Title = ""
            .MB_Button = Error
            Call .MBOX
       End With
       iH1_Gyomu.SetFocus
       Exit Sub
    End If

    CT0040_FileCd = iH1_Gyomu
    CT0040_TEXTTYPE = 0
    CT0040_TEXTName = ""
    
    For Each gObj In Op_TnsType
        If gObj.Value And Trim(gObj.Caption) <> "" Then
           CT0040_TEXTTYPE = gObj.Index
           CT0040_TEXTName = Replace(gObj.Caption, " ", "")
           Exit For
        End If
    Next gObj
    If CT0040_TEXTName = "" Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "   ＴＥＸＴ変換種類を選択して下さい。     "
            .MB_Title = ""
            .MB_Button = Error
            Call .MBOX
       End With
       Exit Sub
    End If
    
'+------------------------------------------------------------ 2007/03/05 Delete
'    If Trim(iH1_GyoNm) = "" And CT0040_TEXTTYPE < 2 Then
'+------------------------------------------------------------ 2007/03/05 Delete End
'+------------------------------------------------------------ 2007/03/05 Insert
    If Trim(iH1_GyoNm) = "" And CT0040_TEXTTYPE < 3 Then
'+------------------------------------------------------------ 2007/03/05 Insert End
       iH1_GyoNm = CT0040_TEXTName & "(" & iH1_Gyomu & ")"
    End If
    CT0040_GyoName = iH1_GyoNm
    
'+------------------------------------------------------------ 2007/03/05 Delete
'    If CT0040_TEXTTYPE = 1 Then
'+------------------------------------------------------------ 2007/03/05 Delete End
'+------------------------------------------------------------ 2007/03/05 Insert
    If CT0040_TEXTTYPE > 0 Then
'+------------------------------------------------------------ 2007/03/05 Insert End
        CCT0046.Show vbModal
        Unload CCT0046
        Set CCT0046 = Nothing
    Else
        CCT0045.Show vbModal
        Unload CCT0045
        Set CCT0045 = Nothing
    End If
    If CT0040_RTcd Then
       Call PB_CAN_Click
    End If
End Sub



