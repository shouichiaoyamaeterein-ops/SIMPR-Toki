VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Begin VB.Form CXU0935 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "売上実績取消"
   ClientHeight    =   1785
   ClientLeft      =   6150
   ClientTop       =   5280
   ClientWidth     =   6030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin CisYM_V60.CisYM iH1_YM 
      Height          =   420
      Left            =   2130
      TabIndex        =   2
      Top             =   900
      Width           =   1380
      _ExtentX        =   2328
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cSize           =   -1  'True
      cChkResult      =   0   'False
      cAutoEject      =   0   'False
   End
   Begin VB.PictureBox PNL_PF 
      Height          =   705
      Left            =   3960
      ScaleHeight     =   645
      ScaleWidth      =   1875
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   1935
      Begin CisBtn_60.CisBtn PB_Ent 
         Height          =   585
         Left            =   990
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "選択"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   128
         BtnShadowColor  =   16777215
         PFCaption       =   "(Enter)"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   16777215
         PFShadowColor   =   0
         PFBackColor     =   128
      End
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   90
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "戻る"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   32768
         BtnShadowColor  =   16777215
         PFCaption       =   "( Esc )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   49152
      End
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 売上実績取消 】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   1
      cPositionX      =   30
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   4860
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         BackColor       =   14737632
         ForeColor       =   16711680
         Caption         =   "SIMPR-A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8454143
         cFont3DColor2   =   16777215
         cFont3D         =   1
         cAlingnment     =   7
         cBoderStyle     =   2
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "実績取消年月"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   450
      TabIndex        =   0
      Top             =   990
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  '不透明
      Height          =   1065
      Left            =   150
      Shape           =   4  '丸みのある長方形
      Top             =   570
      Width           =   3675
   End
End
Attribute VB_Name = "CXU0935"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   売上実績取消
'**       フォームID    :   CXU0935
'**       処理概要      :
'**
'**       作  成  日    :   2005/06/09  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim FormCap         As String
    
    Dim CisFrm          As CISFormContorl
    
    Dim mCount          As Long
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
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    If RV_Left <> 0 Or RV_Top <> 0 Then
        Me.Move RV_Left, RV_Top
    Else
        Call CisFrm.Frm_Center(False, x1024y768)
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormCap = Me.Caption
    
    RV_Rtn = False
    
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
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
'                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
'                                  Dummy.Enabled = False
                               End If
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Call PB_Can_Click
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_Can_Click()
    Unload CXU0935
End Sub
'+----------------------------+
'+  売上実績取消処理
'+----------------------------+
Private Sub ReturnPress()
    Dim wJisseki    As String
        
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　実績取消年月を入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If
    
    If Not iH1_YM.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　正しい年月にて入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If
    
    ' 売上実績存在確認
    gSL_Select = "Select Count(*) 件数 From  売上実績テーブル "
    gSL_Select = gSL_Select & " Where 実績年月 = '" & iH1_YM & "'"
    If Not UJTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　売上実績処理未処理です。           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If
    If UJT.件数 = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　売上実績処理未処理です。           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If
    
    ' 処理確認メッセージ
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     " & StrConv(iH1_YM.cYear, vbWide) & " 年" & _
                               StrConv(CisFun.RSetFld(iH1_YM.cMonth, 2, "#0"), vbWide) & " 月度 " & _
                               "売上実績取消処理を行います。          "
        .MB_Title = "実績取消"
        .MB_MSG(4) = "                    よろしいですか？                 "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Sub
    End With
'
' 売上処理対象ファイルを求める  1=出荷テーブル,1≠検収テーブル
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 50
    Call IRNRead("", 1)
    wJisseki = Mid(IRN.oyian, 1, 1)
    
    
    With CisDB
        On Error GoTo ReturnPress_er
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

        Call ItemsClearUHR
        
        With UHR
            .売上年月 = iH1_YM
            .処理区分 = "D"
'            .処理件数 = KNT.件数
            .作成元区分 = wJisseki
            .作成者 = gTanto
        End With
        Call UHRInsert
' ( 実績削除処理 )
        .SQL = "DELETE FROM 売上実績テーブル "
        .SQL = .SQL & " WHERE 実績年月 = '" & iH1_YM & "'"
        If Not .DBExec Then GoTo ReturnPress_er

        .SQL = "DELETE FROM 売上明細テーブル "
        .SQL = .SQL & " WHERE 実績年月 = '" & iH1_YM & "'"
        If Not .DBExec Then GoTo ReturnPress_er
'
' ( 実績取消処理 )
        If wJisseki = "1" Then
           .SQL = "Update 出荷テーブル Set "
           .SQL = .SQL & " 実績年月 = '' "
           .SQL = .SQL & " WHERE 実績年月 = '" & iH1_YM & "'"
        Else
           .SQL = "Update 検収テーブル Set "
           .SQL = .SQL & " 実績年月 = '' "
           .SQL = .SQL & " WHERE 実績年月 = '" & iH1_YM & "'"
        End If
        If Not .DBExec Then GoTo ReturnPress_er
        
        .SQL = "Update 型費テーブル Set "
        .SQL = .SQL & " 実績年月 = '' "
        .SQL = .SQL & " WHERE 実績年月 = '" & iH1_YM & "'"
        If Not .DBExec Then GoTo ReturnPress_er
        
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

        On Error GoTo 0
    End With
    
    RV_Rtn = True
    Call PB_Can_Click
    Exit Sub
ReturnPress_er:
    Call CisFun.ErrorBox
    End
End Sub
