VERSION 5.00
Begin VB.Form CXHZ0110 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "材料発注処理"
   ClientHeight    =   5175
   ClientLeft      =   3045
   ClientTop       =   2685
   ClientWidth     =   5790
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
   ScaleHeight     =   5175
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      Left            =   4605
      TabIndex        =   9
      Top             =   3300
      Width           =   705
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      Left            =   4605
      TabIndex        =   11
      Top             =   3825
      Width           =   705
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   7
      Left            =   4605
      TabIndex        =   13
      Top             =   4350
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　臨時かんばん　発行"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   10
      Left            =   900
      TabIndex        =   17
      Top             =   3300
      Width           =   3285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　現品票　発行"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   9
      Left            =   900
      TabIndex        =   16
      Top             =   3825
      Width           =   3285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　端数ラベル　発行"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   8
      Left            =   900
      TabIndex        =   15
      Top             =   4350
      Width           =   3285
   End
   Begin VB.Shape Back_SyoriNM 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   5
      Left            =   705
      Shape           =   4  '丸みのある長方形
      Top             =   3165
      Width           =   3690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　納番　採番"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   7
      Left            =   915
      TabIndex        =   14
      Top             =   3330
      Width           =   3285
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   7
      Left            =   4410
      Shape           =   4  '丸みのある長方形
      Top             =   3165
      Width           =   1065
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   5
      Left            =   225
      Stretch         =   -1  'True
      Top             =   3165
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Back_SyoriNM 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   6
      Left            =   720
      Shape           =   4  '丸みのある長方形
      Top             =   3690
      Width           =   3690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　納品書　発行"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   6
      Left            =   930
      TabIndex        =   12
      Top             =   3855
      Width           =   3285
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   6
      Left            =   4410
      Shape           =   4  '丸みのある長方形
      Top             =   3690
      Width           =   1065
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   6
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3690
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Back_SyoriNM 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   7
      Left            =   720
      Shape           =   4  '丸みのある長方形
      Top             =   4215
      Width           =   3690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　 支給書　発行"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   5
      Left            =   930
      TabIndex        =   10
      Top             =   4380
      Width           =   3285
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   5
      Left            =   4410
      Shape           =   4  '丸みのある長方形
      Top             =   4215
      Width           =   1065
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   7
      Left            =   240
      Stretch         =   -1  'True
      Top             =   4215
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img_No 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   480
      Picture         =   "CXHZ0110.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img_End 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   480
      Left            =   0
      Picture         =   "CXHZ0110.frx":030A
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Title 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "【　材料発注処理　】"
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
      Left            =   1305
      TabIndex        =   8
      Top             =   270
      Width           =   3000
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  '不透明
      Height          =   660
      Left            =   135
      Shape           =   4  '丸みのある長方形
      Top             =   60
      Width           =   5565
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   4605
      TabIndex        =   7
      Top             =   2760
      Width           =   705
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   4
      Left            =   225
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   4
      Left            =   4395
      Shape           =   4  '丸みのある長方形
      Top             =   2640
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　発注明細書　発行"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   4
      Left            =   900
      TabIndex        =   6
      Top             =   2790
      Width           =   3285
   End
   Begin VB.Shape Back_SyoriNM 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   4
      Left            =   705
      Shape           =   4  '丸みのある長方形
      Top             =   2640
      Width           =   3690
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   4590
      TabIndex        =   5
      Top             =   2250
      Width           =   705
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   3
      Left            =   210
      Stretch         =   -1  'True
      Top             =   2115
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   3
      Left            =   4380
      Shape           =   4  '丸みのある長方形
      Top             =   2115
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　自社伝票　ファイル出力"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   885
      TabIndex        =   4
      Top             =   2280
      Width           =   3285
   End
   Begin VB.Shape Back_SyoriNM 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   3
      Left            =   690
      Shape           =   4  '丸みのある長方形
      Top             =   2115
      Width           =   3690
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   4590
      TabIndex        =   3
      Top             =   1725
      Width           =   705
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   2
      Left            =   195
      Stretch         =   -1  'True
      Top             =   1590
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   2
      Left            =   4365
      Shape           =   4  '丸みのある長方形
      Top             =   1590
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　自社伝票　発行"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   885
      TabIndex        =   2
      Top             =   1755
      Width           =   3285
   End
   Begin VB.Shape Back_SyoriNM 
      BackColor       =   &H00404040&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   2
      Left            =   675
      Shape           =   4  '丸みのある長方形
      Top             =   1590
      Width           =   3690
   End
   Begin VB.Label lb_Cnt 
      Alignment       =   1  '右揃え
      BackStyle       =   0  '透明
      Caption         =   "Z,ZZZ"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   4575
      TabIndex        =   1
      Top             =   1155
      Width           =   705
   End
   Begin VB.Image img_Ok 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   510
      Index           =   1
      Left            =   195
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
      Left            =   4350
      Shape           =   4  '丸みのある長方形
      Top             =   1020
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "　納番採番・発注累積 件数"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   870
      TabIndex        =   0
      Top             =   1185
      Width           =   3285
   End
   Begin VB.Shape Back_SyoriNM 
      BackColor       =   &H00000000&
      BackStyle       =   1  '不透明
      Height          =   510
      Index           =   1
      Left            =   660
      Shape           =   4  '丸みのある長方形
      Top             =   1020
      Width           =   3690
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '不透明
      Height          =   4320
      Left            =   75
      Shape           =   4  '丸みのある長方形
      Top             =   735
      Width           =   5610
   End
End
Attribute VB_Name = "CXHZ0110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  材料発注処理
'**       フォームID    :  CXHZ0110
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/21  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim CisFrm          As CISFormContorl
    Dim SyoriKb         As String
    
    Dim SyoriCColor     As Long
    Dim SyoriBColor     As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    
    For gInt = 1 To 7
        lb_Cnt(gInt).Caption = ""
    Next
    Select Case RTrim(Command)
        Case "G": Title = "【　材料発注処理　】《発注画面入力》"
        Case "K": Title = "【　材料発注処理　】《かんばん読取》"
        Case "J": Title = "【　材料受領処理　】《入荷実績入力》"
    End Select
    SyoriKb = RTrim(Command)
    SyoriCColor = &HFF8080
    SyoriBColor = &H0&
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
'    Call CisFrm.ExpressionForm(100, 600)
    
    FormAct = Not FormAct
    
    Call ReturnPress
    
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
'+---------------------+
'+   ﾃﾞｰﾀﾍﾞｰｽ切断
'+---------------------+
    Call CisDB.DBDISConnect
'+---------------------+
'+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
'+---------------------+
    Call CisFrm.UnLoadDisp
    End
End Sub
'
'+----------------------------+
'+       処  理   開  始       +
'+----------------------------+
Private Sub ReturnPress()
    DoEvents
'
'--------------------------------------------------'
'   納番・発注管理NO　採番、発注累積
'--------------------------------------------------'
    Back_SyoriNM(1).BackColor = SyoriCColor
    DoEvents
    If Not Update_Proc Then Exit Sub
        
    Back_SyoriNM(1).BackColor = SyoriBColor
    img_Ok(1).Visible = True
    DoEvents
    
'--------------------------------------------------'
'   (( JAMA ))ファイル出力      材料はなし
'--------------------------------------------------'
'''    If CisFun.Val2(lb_Cnt(1)) <> 0 Then
'''        Back_SyoriNM(1).BackColor = SyoriCColor
'''        DoEvents
'''        If Not CXHZ0370TFile(0) Then Exit Sub
'''        Back_SyoriNM(1).BackColor = SyoriBColor
'''        img_Ok(1).Visible = True
'''        DoEvents
'''    End If
'--------------------------------------------------'
'   ((自社伝票))納品書・受領書・支給書 発行
'--------------------------------------------------'
'    lb_Cnt(2) = "1"
    If CisFun.Val2(lb_Cnt(2)) <> 0 Then
        Back_SyoriNM(2).BackColor = SyoriCColor
        DoEvents
        If Not CXHZ0310LPrint(0) Then Exit Sub
        Back_SyoriNM(2).BackColor = SyoriBColor
        img_Ok(2).Visible = True
        DoEvents
    End If
'--------------------------------------------------'
'   ((自社伝票))納品書・受領書・支給書 ファイル出力
'--------------------------------------------------'
'    lb_Cnt(3) = "1"
    If CisFun.Val2(lb_Cnt(3)) <> 0 Then
        Back_SyoriNM(3).BackColor = SyoriCColor
        If Not CXHZ0320TFile(0) Then Exit Sub
        Back_SyoriNM(3).BackColor = SyoriBColor
        img_Ok(3).Visible = True
        DoEvents
    End If
'--------------------------------------------------'
'   明細書 発行
'--------------------------------------------------'
'    lb_Cnt(4) = "1"
    If CisFun.Val2(lb_Cnt(4)) <> 0 Then
        Back_SyoriNM(4).BackColor = SyoriCColor
        DoEvents
        If Not CXHZ0330LPrint(0, gDemoPrt) Then Exit Sub
        Back_SyoriNM(4).BackColor = SyoriBColor
        img_Ok(4).Visible = True
        DoEvents
    End If
'--------------------------------------------------'
'   臨時かんばん 発行
'--------------------------------------------------'
'    lb_Cnt(5) = "1"
    If CisFun.Val2(lb_Cnt(5)) <> 0 Then
        Back_SyoriNM(5).BackColor = SyoriCColor
        DoEvents
        If Not CXHZ0340LPrint(0, gDemoPrt) Then Exit Sub
        Back_SyoriNM(5).BackColor = SyoriBColor
        img_Ok(5).Visible = True
        DoEvents
    End If
'--------------------------------------------------'
'   現品票 発行
'--------------------------------------------------'
'    lb_Cnt(6) = "1"
    If CisFun.Val2(lb_Cnt(6)) <> 0 Then
        Back_SyoriNM(6).BackColor = SyoriCColor
        DoEvents
        If Not CXHZ0350LPrint(0, gDemoPrt) Then Exit Sub
        Back_SyoriNM(6).BackColor = SyoriBColor
        img_Ok(6).Visible = True
        DoEvents
    End If
'--------------------------------------------------'
'   端数ラベル 発行
'--------------------------------------------------'
'    lb_Cnt(7) = "1"
    If CisFun.Val2(lb_Cnt(7)) <> 0 Then
        Back_SyoriNM(7).BackColor = SyoriCColor
        DoEvents
        If Not CXHZ0360LPrint(0, gDemoPrt) Then Exit Sub
        Back_SyoriNM(7).BackColor = SyoriBColor
        img_Ok(7).Visible = True
        DoEvents
    End If

'--------------------------------------------------'
'   処理終了
'--------------------------------------------------'
   Call EndProc
End Sub
'+----------------------------+
'+    納入番号採番・発注累積    +
'+----------------------------+
Private Function Update_Proc()
    Update_Proc = False
    Dim TCnt(9)     As Long
    Dim LogText     As String
    
    With CisFun
        .Log_PutFlg = True
        .Log_FileInitial = "HCZ"
        .Log_FilePath = gIniExe & "Log\"
        .Log_DataHD = RTrim(SyoriKb)
        .Log_DataHDTime = あり
        .Log_DelDay = 10
    End With
    Call CisFun.LogPut(Title & "【開始】 処理端末(" & CisFun.WSID & ")")
'発注打切更新
    With CisDB
        If SyoriKb = "J" Then
            .SQL = "材料発注入荷処理"
        Else
            .SQL = "材料発注処理"
        End If
        .StoadoCount = 14
        .ParaNo = 0: .ParaIO = Return用
        If SyoriKb = "J" Then
            .ParaNo = 1: .ParaIO = OutPut用            '受領件数
        Else
            .ParaNo = 1: .ParaIO = Input用             '処理区分
        End If
        .ParaNo = 2: .ParaIO = OutPut用            '入力件数/分納件数
        .ParaNo = 3: .ParaIO = OutPut用            '発注累積件数
        .ParaNo = 4: .ParaIO = OutPut用            '納入番号採番件数
        .ParaNo = 5: .ParaIO = OutPut用            'JAMA出力件数
        .ParaNo = 6: .ParaIO = OutPut用            '納品書頁数
        .ParaNo = 7: .ParaIO = OutPut用            '支給書頁数
        .ParaNo = 8: .ParaIO = OutPut用            '受領書頁数
        .ParaNo = 9: .ParaIO = OutPut用            '明細書納入番号件数
        .ParaNo = 10: .ParaIO = OutPut用            '臨時かんばん件数
        .ParaNo = 11: .ParaIO = OutPut用            '現品票件数
        .ParaNo = 12: .ParaIO = OutPut用            '端数ラベル件数
        .ParaNo = 13: .ParaIO = OutPut用            '自社伝票ファイル出力件数
        .ParaNo = 14: .ParaIO = OutPut用            'エラーコード
'
        If SyoriKb <> "J" Then
            .ParaNo = 1: .ParaValue = SyoriKb
        End If
'
        If Not .DBStored Then
            MsgBox "ストアドエラー "
            GoTo Update_Proc_Err
        Else
            .ParaNo = 14
            If .ParaValue <> 0 Then GoTo Update_Proc_Err
            
            If SyoriKb = "J" Then
                .ParaNo = 1
                LogText = "受領件数 = " & .ParaValue & ", "
                .ParaNo = 2
                LogText = LogText & "分納件数 = " & .ParaValue & " "
            Else
                .ParaNo = 2
                LogText = "発注入力件数 = " & .ParaValue & " "
            End If
            
            .ParaNo = 3
            lb_Cnt(1) = Format(.ParaValue, "#,###")
            Set img_Ok(1).Picture = img_End.Picture: DoEvents
            LogText = LogText & "発注累積件数 = " & .ParaValue & " "
            
            .ParaNo = 4
            LogText = "　　　　　" & LogText & "納入番号件数 = " & .ParaValue & " "
            Call CisFun.LogPut(LogText)
            
            LogText = "　　　　　"
            For gInt = 5 To 13
                .ParaNo = gInt
                TCnt(gInt - 4) = .ParaValue
                Select Case gInt
                    Case 5: gStr = "JAMA出力頁数 = "
                    Case 6: gStr = "自社出力頁数 = "
                    Case 7: gStr = "納品書頁数 = "
                    Case 8: gStr = "受領書頁数 = "
                    Case 9: gStr = "支給書頁数 = "
                    Case 10: gStr = "明細書頁数 = "
                            Call CisFun.LogPut(LogText)
                            LogText = "　　　　　"
                    Case 11: gStr = "臨時かんばん件数 = "
                    Case 12: gStr = "現品票件数 = "
                    Case 13: gStr = "端数ラベル件数 = "
                End Select
                LogText = LogText & gStr & .ParaValue & "  "
                
            Next
            Call CisFun.LogPut(LogText)
            
            For gCnt1 = 2 To 7
                Select Case gCnt1
                    Case 1: gCnt2 = TCnt(1)
                    Case 2: gCnt2 = TCnt(3) + TCnt(4) + TCnt(5)
                    Case 3: gCnt2 = TCnt(2)
                    Case 4: gCnt2 = TCnt(6)
                    Case 5: gCnt2 = TCnt(7)
                    Case 6: gCnt2 = TCnt(8)
                    Case 7: gCnt2 = TCnt(9)
                End Select
                lb_Cnt(gCnt1) = Format(gCnt2, "#,###")
                If gCnt2 = 0 Then
                    Set img_Ok(gCnt1).Picture = img_No.Picture
                    img_Ok(gCnt1).Visible = True
                Else
                    Set img_Ok(gCnt1).Picture = img_End.Picture
                End If
            Next
        End If
        DoEvents
    End With
    
    Call CisFun.LogPut(Title & "【正常終了】　")
    CisFun.Log_PutFlg = False
    Call CisFun.LogPut("")
    
    Update_Proc = True
    Exit Function
Update_Proc_Err:
    Select Case CisDB.ParaValue
'発注処理
        Case 1: gStr = "    【　材料発注累積テーブル　削除エラー！】"
        Case 2: gStr = "    【　材料発注累積ワーク　削除エラー！】"
        Case 3: gStr = "    【　帳票出力テーブル(25)　削除エラー！】"
        Case 4: gStr = "    【　帳票出力テーブル(26)　削除エラー！】"
        Case 5: gStr = "    【　帳票出力テーブル(27)　削除エラー！】"
        Case 6: gStr = "    【　帳票出力テーブル(28)　削除エラー！】"
        Case 7: gStr = "    【　帳票出力テーブル(29)　削除エラー！】"
        Case 8: gStr = "    【　帳票出力テーブル(30)　削除エラー！】"
        Case 9: gStr = "    【　帳票出力テーブル(31)　削除エラー！】"
        
        Case 11: gStr = "    【　材料発注累積ワーク(G)　出力エラー！】"
        Case 12: gStr = "    【　材料発注累積ワーク(K)　出力エラー！】"
        Case 13: gStr = "    【　材料発注累積ワーク(K)　更新エラー！】"
        
        Case 20: gStr = "    【　納品書管理NO　採番エラー！】"
        Case 21: gStr = "    【　納入番号　採番エラー！】"
        Case 22: gStr = "    【　材料発注管理NO　採番エラー！】"
        Case 23: gStr = "    【　材料発注累積テーブル　出力エラー！】"
        Case 24: gStr = "    【　材料発注テーブル　出力エラー！】"
        Case 25: gStr = "    【　材料発注画面ワーク　削除エラー！】"
        Case 26: gStr = "    【　材料発注かんばん予定テーブル　削除エラー！】"
        Case 31: gStr = "    【　帳票出力テーブル(128)　出力エラー！】"
        Case 32: gStr = "    【　帳票出力テーブル(29)　出力エラー！】"
        Case 33: gStr = "    【　帳票出力テーブル(30)　出力エラー！】"
        Case 34: gStr = "    【　帳票出力テーブル(31)　出力エラー！】"
        
        Case 71: gStr = "    【　材料発注明細書№　採番エラー！】"
        Case 72: gStr = "    【　帳票出力テーブル(28)　出力エラー！】"
        Case 73: gStr = "    【　材料発注テーブル　明細NO更新エラー！】"

'発注伝票発行データ作成
        Case 35: gStr = "    【　帳票出力テーブル(25)　出力エラー！】"
        Case 36: gStr = "    【　帳票出力テーブル(26)　出力エラー！】"
        Case 37: gStr = "    【　帳票出力テーブル(27)　出力エラー！】"
        
        Case 41: gStr = "    【　総頁数・物理頁数　更新エラー！】"
        Case 42: gStr = "    【　END表示区分(納品書)　更新エラー！】"
        Case 43: gStr = "    【　END表示区分(受領書)　更新エラー！】"
        Case 44: gStr = "    【　END表示区分(支給書)　更新エラー！】"
        Case 45: gStr = "    【　材料発注帳票発行テーブル(納品書)　出力エラー！】"
        Case 46: gStr = "    【　材料発注帳票発行テーブル(受領書)　出力エラー！】"
        Case 47: gStr = "    【　材料発注帳票発行テーブル(支給書)　出力エラー！】"
        Case 48: gStr = "    【　材料発注帳票発行テーブル　更新エラー！】"
        Case 49: gStr = "    【　材料発注累積テーブル　更新エラー！】"
        Case 50: gStr = "    【　総頁数・物理頁数　更新エラー！】"
        Case 51: gStr = "    【　END表示区分(納品書)　更新エラー！】"
        Case 52: gStr = "    【　END表示区分(受領書)　更新エラー！】"
        Case 53: gStr = "    【　END表示区分(支給書)　更新エラー！】"
        
        Case 61: gStr = "    【　ALLEND表示区分　更新エラー！】"
        Case 62: gStr = "    【　物理総頁数　更新エラー！】"
        Case 63: gStr = "    【　ALLEND表示区分　更新エラー！】"
        Case 64: gStr = "    【　物理総頁数　更新エラー！】"
        Case 65: gStr = "    【　処理端末　更新エラー！】"
        Case 66: gStr = "    【　材料発注累積テーブル　更新エラー！】"
        Case 67: gStr = "    【　材料発注累積テーブル　削除エラー！】"

'発注入荷処理
        Case 101: gStr = "    【　材料発注累積テーブル　削除エラー！】"
        Case 102: gStr = "    【　材料発注累積ワーク　削除エラー！】"
        Case 103: gStr = "    【　帳票出力テーブル(25)　削除エラー！】"
        Case 104: gStr = "    【　帳票出力テーブル(26)　削除エラー！】"
        Case 105: gStr = "    【　帳票出力テーブル(27)　削除エラー！】"
        Case 106: gStr = "    【　帳票出力テーブル(28)　削除エラー！】"
        Case 107: gStr = "    【　帳票出力テーブル(29)　削除エラー！】"
        Case 108: gStr = "    【　帳票出力テーブル(30)　削除エラー！】"
        Case 109: gStr = "    【　帳票出力テーブル(31)　削除エラー！】"
        
        Case 114: gStr = "    【　材料発注累積テーブル(J受領)　出力エラー！】"
        Case 115: gStr = "    【　材料発注累積ワーク(J分納)　出力エラー！】"
        Case 116: gStr = "    【　材料入荷実績処理ワーク　更新エラー！】"
        
        Case 120: gStr = "    【　材料納品書管理NO　採番エラー！】"
        Case 122: gStr = "    【　材料発注管理NO　採番エラー！】"
        Case 123: gStr = "    【　材料発注累積テーブル　出力エラー！】"
        Case 124: gStr = "    【　材料発注テーブル　出力エラー！】"
        Case 125: gStr = "    【　材料発注画面ワーク　削除エラー！】"
        Case 126: gStr = "    【　材料発注かんばん予定テーブル　削除エラー！】"
        Case 131: gStr = "    【　帳票出力テーブル(128)　出力エラー！】"
        Case 132: gStr = "    【　帳票出力テーブル(29)　出力エラー！】"
        Case 133: gStr = "    【　帳票出力テーブル(30)　出力エラー！】"
        Case 134: gStr = "    【　帳票出力テーブル(31)　出力エラー！】"
        
        Case 171: gStr = "    【　材料発注明細書№　採番エラー！】"
        Case 172: gStr = "    【　帳票出力テーブル(28)　出力エラー！】"
        Case 173: gStr = "    【　材料発注テーブル　明細NO更新エラー！】"
        
        Case Else: gStr = "    【　その他　エラー！】"
    End Select
    gStr = gStr & "  Code = " & CisDB.ParaValue & "     "
    MsgBox gStr
    Call CisFun.LogPut(gStr)
    Call CisFun.LogPut(Title & "【異常終了】　")
    Call EndProc
End Function


