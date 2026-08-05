VERSION 5.00
Begin VB.Form CKHZ0900 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "材料発注かんばん後方処理"
   ClientHeight    =   3615
   ClientLeft      =   4425
   ClientTop       =   3345
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.Image img_No 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   480
      Picture         =   "CKHZ0900.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img_End 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   0
      Picture         =   "CKHZ0900.frx":030A
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "材料発注かんばん リスト作成･更新処理"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   195
      TabIndex        =   10
      Top             =   270
      Width           =   5625
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   660
      Left            =   90
      Shape           =   4  '丸みのある長方形
      Top             =   60
      Width           =   5910
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   4965
      TabIndex        =   9
      Top             =   2955
      Width           =   705
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   4
      Left            =   225
      Stretch         =   -1  'True
      Top             =   2835
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   4
      Left            =   4740
      Shape           =   4  '丸みのある長方形
      Top             =   2835
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "材料発注予定データ累積"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   4
      Left            =   915
      TabIndex        =   8
      Top             =   2985
      Width           =   3630
   End
   Begin VB.Shape ProcName 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   4
      Left            =   705
      Shape           =   4  '丸みのある長方形
      Top             =   2835
      Width           =   4035
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   4950
      TabIndex        =   7
      Top             =   2220
      Width           =   705
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   3
      Left            =   210
      Stretch         =   -1  'True
      Top             =   2085
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   3
      Left            =   4725
      Shape           =   4  '丸みのある長方形
      Top             =   2085
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "材料発注打切リスト"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   900
      TabIndex        =   6
      Top             =   2250
      Width           =   3630
   End
   Begin VB.Shape ProcName 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   3
      Left            =   690
      Shape           =   4  '丸みのある長方形
      Top             =   2085
      Width           =   4035
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   4935
      TabIndex        =   5
      Top             =   1695
      Width           =   705
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   2
      Left            =   195
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   2
      Left            =   4710
      Shape           =   4  '丸みのある長方形
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "材料発注打切警告リスト"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   885
      TabIndex        =   4
      Top             =   1725
      Width           =   3630
   End
   Begin VB.Shape ProcName 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   2
      Left            =   675
      Shape           =   4  '丸みのある長方形
      Top             =   1560
      Width           =   4035
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   4635
      TabIndex        =   3
      Top             =   3765
      Width           =   705
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   1
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   1
      Left            =   4410
      Shape           =   4  '丸みのある長方形
      Top             =   3630
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "平準化エラーリスト"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   930
      TabIndex        =   2
      Top             =   3795
      Width           =   3285
   End
   Begin VB.Shape ProcName 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   1
      Left            =   720
      Shape           =   4  '丸みのある長方形
      Top             =   3630
      Width           =   3690
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   4920
      TabIndex        =   1
      Top             =   1155
      Width           =   705
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   0
      Left            =   180
      Stretch         =   -1  'True
      Top             =   1020
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   0
      Left            =   4695
      Shape           =   4  '丸みのある長方形
      Top             =   1020
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "材料かんばん読取エラーリスト"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   870
      TabIndex        =   0
      Top             =   1185
      Width           =   3630
   End
   Begin VB.Shape ProcName 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   0
      Left            =   660
      Shape           =   4  '丸みのある長方形
      Top             =   1020
      Width           =   4035
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  '不透明
      Height          =   2760
      Left            =   90
      Shape           =   4  '丸みのある長方形
      Top             =   780
      Width           =   5865
   End
End
Attribute VB_Name = "CKHZ0900"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  材料発注かんばん後方処理
'**       フォームID    :  CKNZ0900
'**       処理概要      :
'**
'**       作  成  日    :  2004/01/21  By CIS
'**       変  更  日    :  2005/11/07  By CIS オプション追加（発注かんばん読取ｴﾗｰﾘｽﾄ発行・発注打切リスト発行）
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim CisFrm                  As CISFormContorl
    
    Dim mProcColor      As Long
    Dim mEndColor       As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    
    kBas_ReadOk = False
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me

    mProcColor = &HFF0000
    mEndColor = ProcName(0).BackColor
    
    For Each gObj In lb_Cnt
        If gObj.Index < 4 Then
            gSL_Select = "SELECT count(*) 件数 FROM 帳票出力テーブル CH "
            gSL_Select = gSL_Select & "  Left Outer Join 材料かんばん読取エラーテーブル KE"
            gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
            gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分"
            gSL_Select = gSL_Select & " Where 端末番号 = Host_Name()"
            gSL_Select = gSL_Select & "   And 帳票区分 = " & gObj.Index + 21
            If gObj.Index = 0 Then
               gSL_Select = gSL_Select & "   And ISNULL(KE.材料管理番号,0) <> 0"
            End If
            Call CHTRead(gSL_Select, 1)
        Else
            gSL_Select = "SELECT count(*) 件数 FROM 材料発注かんばん枚数累積テーブル"
            gSL_Select = gSL_Select & " Where 処理端末 = Host_Name()"
            Call CHTRead(gSL_Select, 1)
        End If
        gObj.Caption = Format(CHT.件数, "#,###")
        If CHT.件数 = 0 Then
            Set img_Ok(gObj.Index).Picture = img_No.Picture
            img_Ok(gObj.Index).Visible = True
        Else
            Set img_Ok(gObj.Index).Picture = img_End.Picture
        End If
    Next gObj
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(False)
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'    Call CisFrm.ExpressionForm(10, 10)
    
    Call ReturnPress
    
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
    
''   ( 続行(Enter) )
'    If PB_OK.Visible Then
'       If Key_Code = vbKeyReturn Then
'           If IVENT = W_IVENT Then
'               Call PB_OK_Click
'           End If
'           Key_Acc = True: Exit Function
'       End If
'    End If
''   ( 取消(Escape) )
'    If PB_CAN.Visible Then
'       If Key_Code = vbKeyEscape Then
'           If IVENT = W_IVENT Then
'               Call PB_CAN_Click
'           End If
'           Key_Acc = True: Exit Function
'       End If
'    End If
    
End Function
''+----------------------------+
''+    続  行  キ ー(Enter)    +
''+----------------------------+
'Private Sub PB_OK_Click()
'    kBas_ReadOk = True
'
'    Call EndProc
'End Sub
''+----------------------------+
''+    取  消  キ ー(Escape)    +
''+----------------------------+
'Private Sub PB_CAN_Click()
'    kBas_ReadOk = False
'
'    Call EndProc
'End Sub
'+----------------------------+
'+       終  了  処  理        +
'+----------------------------+
Private Sub EndProc()
    Me.Hide
End Sub
'
'+----------------------------+
'+       処  理   開  始       +
'+----------------------------+
Private Sub ReturnPress()
    DoEvents
    
    ' 材料かんばん読取エラーリスト
    If CisFun.Val2(lb_Cnt(0)) <> 0 And kBas_Check.OpHErrPrt = 0 Then                '2005/11/07追加
        ProcName(0).BackColor = mProcColor
        DoEvents
        
        If Not CKHZ0910LPrint Then Exit Sub
        img_Ok(0).Visible = True
        ProcName(0).BackColor = mEndColor
        DoEvents
    End If
    ' 材料発注打切警告リスト
    If CisFun.Val2(lb_Cnt(2)) <> 0 And kBas_Check.OpHUchiPrt = 0 Then                '2005/11/07追加
        ProcName(2).BackColor = mProcColor
        DoEvents
        CHHZ0930_CallType = "C"
        If Not CKHZ0930LPrint Then Exit Sub
        img_Ok(2).Visible = True
        ProcName(2).BackColor = mEndColor
        DoEvents
    End If
    ' 材料発注打切リスト
    If CisFun.Val2(lb_Cnt(3)) <> 0 And kBas_Check.OpHUchiPrt = 0 Then                '2005/11/07追加
        ProcName(3).BackColor = mProcColor
        DoEvents
        CHHZ0940_CallType = "C"
        If Not CKHZ0940LPrint Then Exit Sub
        img_Ok(3).Visible = True
        ProcName(3).BackColor = mEndColor
        DoEvents
    End If
    
    ' 材料発注かんばん予定テーブル　出力
    If CisFun.Val2(lb_Cnt(4)) <> 0 Then
       ProcName(4).BackColor = mProcColor
        DoEvents
       With CisDB
'            Call .DBTran(TransBegin, 1)
            ' 材料かんばん重複チェックテーブル削除
            .SQL = "Delete 材料かんばん重複チェックテーブル "
            If kBas_DelDay2 <> 0 Then
               .SQL = .SQL & "    Where 納入日 < Convert(Varchar(8),Getdate(),112)"
            End If
            .ConnectNo = 1
            .DBExec
'            ' かんばんチェックテーブル削除
'            .SQL = "Delete かんばんチェックテーブル "
'            gStr = Format(Date, "YYYYMMDD")
'            If kBas_DelDay1 <> 0 Then
'                gInt = kBas_DelDay1 * -1
'                With CisFun
'                    .Date_RsFormat = fYMD
'                    .Date_Add Format(Date, "YYYYMMDD"), D, gInt
'                    gStr = .Date_Result
'                End With
'            End If
'            .SQL = .SQL & " Where 納入日 < '" & gStr & "'"
'            .ConnectNo = 1
'            .DBExec
'            ' 平準化設定ｵﾌﾟｼｮﾝが発注ありの時、月完了区分更新
'            If kBas_Check.Op1Value <> "" Then
'                .SQL = "UPDATE かんばんチェックテーブル SET "
'                .SQL = .SQL & "月完了区分 = 1 "
'                .SQL = .SQL & "  FROM 発注かんばん枚数累積テーブル KR"
'                .SQL = .SQL & "    Where かんばんチェックテーブル.納入日 = SUBSTRING(KR.納入日,1,6) + '99'"
'                .SQL = .SQL & "      And かんばんチェックテーブル.取引先 = KR.取引先"
'                .SQL = .SQL & "      And かんばんチェックテーブル.受入    = KR.受入"
'                .SQL = .SQL & "      And かんばんチェックテーブル.品番    = KR.品番"
'                .SQL = .SQL & "      And KR.平準化区分 = 3"
'                .SQL = .SQL & "      And KR.平準化月超過日 <> ''"
'                .ConnectNo = 1
'                .DBExec
'            End If
            .SQL = "Insert Into 材料発注かんばん予定テーブル" & vbCr
            .SQL = .SQL & "( 処理端末,処理種別,発注区分," & vbCr
            .SQL = .SQL & "  仕入先,受入,納入日,納入便,手配区分,SK区分," & vbCr
            .SQL = .SQL & "  材料管理番号,材質,板厚,幅,長さ," & vbCr
            .SQL = .SQL & "  背番号,収容数,枚数,発注数," & vbCr
            .SQL = .SQL & "  サイクル,発注日,発注便,発行日,発行時間," & vbCr
            .SQL = .SQL & "  打切区分,打切種別,打切残数,打切日," & vbCr
            .SQL = .SQL & "  端数区分," & vbCr
            .SQL = .SQL & "  端数,平準化区分,平準化開始日,作成種別," & vbCr
            .SQL = .SQL & "  処理区分," & vbCr
            .SQL = .SQL & "  作成日,作成者,作成端末)" & vbCr
            .SQL = .SQL & " Select" & vbCr
            If kBas_Check.Op2Value = "1" Then
               .SQL = .SQL & "処理端末," & vbCr
            Else
               .SQL = .SQL & "  CASE WHEN 発注区分 = '3' Then 処理端末 ELSE '' END," & vbCr
            End If
            .SQL = .SQL & "  読取処理,発注区分," & vbCr
            .SQL = .SQL & "  取引先,受入,納入日,納入便,手配区分,SK区分," & vbCr
            .SQL = .SQL & "  材料管理番号,材質,板厚,幅,長さ," & vbCr
            .SQL = .SQL & "  背番号,収容数,読取枚数,発注数," & vbCr
            .SQL = .SQL & "  サイクル,発注日,発注便,発行日,発行時間," & vbCr
            .SQL = .SQL & "  打切区分,打切種別,打切残数,打切日," & vbCr
            .SQL = .SQL & "  CASE WHEN 打切端数 <> 0 And 打切種別 <> '9' Then 1 Else 0 End," & vbCr
            .SQL = .SQL & "  打切端数,平準化区分,平準化開始日," & vbCr
            .SQL = .SQL & "  CASE WHEN 読取処理 = '1' Then 'I' Else 'K' End," & vbCr
            If kBas_Check.Op2Value = "1" Then
               .SQL = .SQL & "1," & vbCr
            Else
               .SQL = .SQL & "  CASE WHEN 発注区分 = '3' Then 1 ELSE 0 END," & vbCr
            End If
            .SQL = .SQL & "GETDATE()," & vbCr
            .SQL = .SQL & "'" & RTrim(gTanto) & "'," & vbCr
            .SQL = .SQL & "SubString(Host_Name(),1,20)" & vbCr
            .SQL = .SQL & "    From 材料発注かんばん枚数累積テーブル" & vbCr
            .SQL = .SQL & "    Where 処理端末 = Host_Name()" & vbCr
            .SQL = .SQL & "      And 処理区分 = 0" & vbCr
            .ConnectNo = 1
            If .DBExec Then
                .SQL = "Update 材料発注かんばん枚数累積テーブル Set "
                .SQL = .SQL & " 処理区分 = 1"
                .SQL = .SQL & " Where 処理端末 = Host_Name()"
                .ConnectNo = 1
                .DBExec
                Call .DBTran(TransCommit, 1)
            Else
                Call .DBTran(TransRollback, 1)
            End If

            img_Ok(4).Visible = True
            ProcName(4).BackColor = mEndColor
            DoEvents
       End With
    
    End If
    For gLong = 1 To 30000000
    Next gLong

    Call EndProc
End Sub

