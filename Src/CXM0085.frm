VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Begin VB.Form CXM0085 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "担当者メンテ"
   ClientHeight    =   2490
   ClientLeft      =   4410
   ClientTop       =   3840
   ClientWidth     =   7365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   540
      Width           =   165
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4080
      ScaleHeight     =   660
      ScaleWidth      =   3075
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Width           =   3135
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1350
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   3
         BtnCaption      =   "取消"
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
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   2205
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   2
         BtnCaption      =   "入力"
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
      Begin CisBtn_60.CisBtn PB_Pass 
         Height          =   585
         Left            =   30
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   1
         BtnCaption      =   "ﾊﾟｽﾜｰﾄﾞ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   65535
         BtnShadowColor  =   0
         PFCaption       =   "( F1 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   65535
      End
   End
   Begin VB.PictureBox H1_Area1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   7035
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1260
      Width           =   7095
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   915
         Left            =   4920
         Top             =   75
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "ﾊﾟｽﾜｰﾄﾞﾚﾍﾞﾙ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Begin VB.ComboBox H1Cmb_PLebel 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "CXM0085.frx":0000
            Left            =   300
            List            =   "CXM0085.frx":000D
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   3
            Top             =   420
            Width           =   1680
         End
         Begin CisText_V60.CisText iH1_PLebel 
            Height          =   360
            Left            =   60
            TabIndex        =   2
            Top             =   420
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "X"
            MaxLength       =   1
         End
      End
      Begin Cis3D_v60.CIS3D Back_Tanto 
         Height          =   915
         Left            =   60
         Top             =   75
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "担当者ｺｰﾄﾞ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Begin CisText_V60.CisText iH1_Tanto 
            Height          =   375
            Left            =   90
            TabIndex        =   0
            Top             =   420
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXXXXXXXX"
            MaxLength       =   15
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   915
         Left            =   2250
         Top             =   75
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "担 当 者 名"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Begin CisText_V60.CisText iH1_TanNm 
            Height          =   375
            Left            =   60
            TabIndex        =   1
            Top             =   420
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXX1XXXXXXXXX2"
            MaxLength       =   20
            IMEMode         =   4
         End
      End
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 担当者追加  】"
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
      cPositionX      =   100
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   6450
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
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
End
Attribute VB_Name = "CXM0085"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   担当者管理マスタメンテナンス
'**       フォームID    :   CXM0085
'**       処理概要      :
'**
'**       作  成  日    :   2003/11/17  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ

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
    Call CISFrm.InitFld
'
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    If CisDB.PassType = "PASS2" Then
       PB_Pass.Visible = True
    Else
       PB_Pass.Visible = False
    End If
    
    Set c080_VsGridObj = c080_VsGrid.GridObj
'+---------------------+
'+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
'+---------------------+
    With c080_Conbo
        Set .CombObj = H1Cmb_PLebel
        .ComboSet
    End With
    
    If c080_ProcKB = "MNT" Then
       PGName = "【担当者　修正】"
       With c080_VsGridObj
            iH1_Tanto = .TextMatrix(.Row, c080_VsGrid.FixedGet(1))
            iH1_TanNm = .TextMatrix(.Row, c080_VsGrid.FixedGet(2))
            gStr = Trim(.TextMatrix(.Row, c080_VsGrid.FixedGet(4)))
            iH1_PLebel = gStr
            With c080_Conbo
                .Code = gStr
                .TextGet
            End With
            gStr = Trim(.TextMatrix(.Row, c080_VsGrid.FixedGet(5)))
            c080_Info = Mid(gStr, 16)
            Back_Tanto.Enabled = False
            iH1_Tanto.cLostColor = &HFFC0C0
       End With
    Else
       PGName = "【担当者　追加】"
       PB_Pass.Visible = False
       c080_Info = CisDB.PassWord("")
    End If
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CISFrm.Frm_Center(False, x1024y768)

End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormAct = Not FormAct
'
    H1_Area1.Refresh

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
           Case vbKeyF1:       If PB_Pass.Visible Then Call PB_Pass_Click    '【ﾊﾟｽﾜｰﾄﾞ】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click      '【取消】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
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
    Call PB_CAN_Click
End Sub
'+---------------------+
'+    取 消  キ ー     +
'+---------------------+
Private Sub PB_CAN_Click()
    Me.Hide
End Sub
'+---------------------+
'+    入　力  キ ー     +
'+---------------------+
Private Sub PB_ENT_Click()
'
    If Not InputChk Then Exit Sub
    
    If Not DBPut Then Exit Sub
    
    Call PB_CAN_Click
End Sub
'+---------------------+
'+    パスワード        +
'+---------------------+
Private Sub PB_Pass_Click()
    
    c080_infoCD = Trim(iH1_Tanto)
    
    CXM0087.Show vbModal
    Unload CXM0087
    Set CXM0087 = Nothing
End Sub
'+---------------------------------------------+
'+  iH1_PLebel              パスワードレベル
'+---------------------------------------------+
Private Sub iH1_PLebel_LostFocus()
    With c080_Conbo
        .Code = Trim(iH1_PLebel)
        Call .TextGet
    End With
End Sub
'+---------------------------------------------+
'+  H1Cmb_PLebel           パスワードレベル
'+---------------------------------------------+
Private Sub H1Cmb_PLebel_GotFocus()
    H1Cmb_PLebel.BackColor = gIGotSel
End Sub
Private Sub H1Cmb_PLebel_LostFocus()
    H1Cmb_PLebel.BackColor = gILostSel
End Sub
Private Sub H1Cmb_PLebel_Click()
    With c080_Conbo
        Call .CodeGet
        iH1_PLebel = .Code
    End With
End Sub
'+------------------------------------------------+
'+      入力内容チェック
'+------------------------------------------------+
Private Function InputChk() As Boolean
    InputChk = False
    If Trim(iH1_Tanto) = "" Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(2) = "       担当者コードを入力して下さい            "
            .MB_Button = Error
            Call .MBOX
        End With
       iH1_Tanto.SetFocus
       Exit Function
    End If
    If Trim(iH1_TanNm) = "" Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(2) = "       担当者名を入力して下さい            "
            .MB_Button = Error
            Call .MBOX
        End With
       iH1_TanNm.SetFocus
       Exit Function
    End If
'
' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = "SELECT * FROM Irnaksic "
    gSL_Select = gSL_Select & " WHERE bk = '" & Trim(iH1_Tanto) & "'"
    gSL_Select = gSL_Select & "   AND rno = 99"
    If IRNRead(gSL_Select) Then
       If c080_ProcKB = "ADD" Then
          With CisFun
              .MB_Lines = 3
              .MB_MSG(2) = "       担当者が既に登録されています。            "
              .MB_Button = Error
              Call .MBOX
          End With
          Exit Function
       End If
    Else
       If c080_ProcKB <> "ADD" Then
          With CisFun
              .MB_Lines = 3
              .MB_MSG(2) = "       担当者が未登録です。            "
              .MB_Button = Error
              Call .MBOX
          End With
          Exit Function
       End If
    End If
    
    InputChk = True
End Function
'+------------------------------------------------+
'+      DB更新
'+------------------------------------------------+
Private Function DBPut() As Boolean
    Dim wStr1       As String
    Dim wStr2       As String
    Dim SetRow      As Long
    DBPut = False
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "         更新します。                  "
        .MB_MSG(3) = "       よろしいですか？                "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
    If c080_ProcKB = "MNT" Then
        gStr = c080_VsGridObj.TextMatrix(c080_VsGridObj.Row, c080_VsGrid.FixedGet(5))
        wStr1 = c080_Info
        wStr2 = .PassWord(iH1_PLebel)
        .SQL = "Update Irnaksic  Set "
        .SQL = .SQL & "oyian = '" & wStr2 & wStr1 & "',"
        .SQL = .SQL & "oyian2 = '" & RTrim(iH1_TanNm) & "',"
        .SQL = .SQL & "etadpu = GetDate()"
        .SQL = .SQL & " Where bk = '" & RTrim(iH1_Tanto) & "'"
        .SQL = .SQL & "   And rno = 99"
        .DBExec
        
        With c080_VsGridObj
            .TextMatrix(.Row, c080_VsGrid.FixedGet(2)) = iH1_TanNm
            .TextMatrix(.Row, c080_VsGrid.FixedGet(3)) = Trim(H1Cmb_PLebel.Text)
            .TextMatrix(.Row, c080_VsGrid.FixedGet(4)) = iH1_PLebel
            .TextMatrix(.Row, c080_VsGrid.FixedGet(5)) = wStr1 & wStr2
        End With
    Else
        c087_Rtn = False
        c080_infoCD = RTrim(iH1_Tanto)
        CXM0087.Show vbModal
        Unload CXM0087
        Set CXM0087 = Nothing
        If Not c087_Rtn Then Exit Function
        
        wStr2 = .PassWord(iH1_PLebel)
        wStr1 = c080_Info
        SetRow = 0
        With c080_VsGridObj
            For gLong = 1 To .Rows - 1
                If Trim(.TextMatrix(gLong, c080_VsGrid.FixedGet(1))) = "" Then
                   SetRow = gLong
                   Exit For
                End If
            Next gLong
            If gLong = 0 Then
                .Rows = .Rows + 1
                SetRow = .Rows - 1
            End If
            .TextMatrix(SetRow, c080_VsGrid.FixedGet(1)) = RTrim(iH1_Tanto)
            .TextMatrix(SetRow, c080_VsGrid.FixedGet(2)) = RTrim(iH1_TanNm)
            .TextMatrix(SetRow, c080_VsGrid.FixedGet(3)) = Trim(H1Cmb_PLebel.Text)
            .TextMatrix(SetRow, c080_VsGrid.FixedGet(4)) = Trim(iH1_PLebel)
            .TextMatrix(SetRow, c080_VsGrid.FixedGet(5)) = wStr2 & wStr1
        End With
        .SQL = "Insert Into Irnaksic  ( "
        .SQL = .SQL & "bk,rno,oyian,"
        .SQL = .SQL & "oyian2,etadpu"
        .SQL = .SQL & " ) Values ( "
        .SQL = .SQL & "'" & RTrim(iH1_Tanto) & "',"
        .SQL = .SQL & "99,"
        .SQL = .SQL & "'" & wStr2 & wStr1 & "',"
        .SQL = .SQL & "'" & RTrim(iH1_TanNm) & "',"
        .SQL = .SQL & "Getdate()"
        .SQL = .SQL & " )"
        .DBExec
        
        Call c080_VsGrid.VSSort(True)
        With c080_VsGridObj
            For gLong = 1 To .Rows - 1
                If Trim(.TextMatrix(gLong, c080_VsGrid.FixedGet(1))) = Trim(iH1_Tanto) Then
                   .Row = gLong
                   Exit For
                End If
            Next gLong
        End With
    End If
    End With
    
    DBPut = True
End Function

