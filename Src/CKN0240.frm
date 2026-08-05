VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form CKN0240 
   BackColor       =   &H00000000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "かんばん発行"
   ClientHeight    =   2220
   ClientLeft      =   4575
   ClientTop       =   3120
   ClientWidth     =   5715
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton PB_Can 
      Caption         =   "印刷中止"
      Height          =   435
      Left            =   1920
      TabIndex        =   5
      Top             =   1740
      Width           =   1635
   End
   Begin VB.PictureBox Bar 
      Height          =   285
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   5355
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1350
      Width           =   5415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   1125
      Left            =   150
      ScaleHeight     =   1065
      ScaleWidth      =   5355
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   5415
      Begin VB.Label Pr_Cnt 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00C0FFC0&
         Caption         =   " Z,ZZZ / Z,ZZZ"
         Height          =   255
         Left            =   1620
         TabIndex        =   4
         Top             =   810
         Width           =   1755
      End
      Begin VB.Label PrMsg 
         Alignment       =   2  '中央揃え
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "【  かんばん発行処理中  】"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   780
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label PrMsg 
         Alignment       =   2  '中央揃え
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "【  かんばん発行処理中  】"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   810
         TabIndex        =   2
         Top             =   390
         Width           =   3705
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "CKN0240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばん 発行処理
'**       フォームID    :  CKN0240
'**       処理概要      :
'**
'**       作  成  日    :  2003/12/25
'**       変  更  日    :  2004/09/15 印刷中止機能追加
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim wXOnXOffCNT     As Integer
    
    Dim mKansy          As Integer

    Dim mSTHopa         As Integer
    Dim mEDHopa         As Integer

    Dim Prt_Cansel_Flg  As Boolean

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》            +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   #  項 目 初 期 化  #
'   #------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
    
'   #------------------#
'   # 画面中央位置表示 #
'   #------------------#
    Call CisFrm.Frm_Center(False, x1024y768)

End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    Kn_Rtn = False
    Prt_Cansel_Flg = False
    
    If Not KanPrtSY(True) Then Exit Sub
    
    Qr_Jisya = ""
    Qr_HMoto = ""
    ' 自社コード
    Qr_Jisya = RTrim(gCompany)
    '発行元
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    Call IRNRead("", 1)
    Qr_HMoto = RTrim(Mid(IRN.oyian, 1, 7))
    If Trim(Qr_HMoto) = "" Then Qr_HMoto = "SIMPR-A"
    
    If Kn_Hopa <> 5 Then
        gSL_Select = "Select * From 名称マスタ"
        gSL_Select = gSL_Select & " Where 区分名称 = 'かんばん発行ホッパ指定'"
        gSL_Select = gSL_Select & " And   区分タイプ = ''"
        gSL_Select = gSL_Select & " And   値 = '" & Format(Kn_Hopa, "0") & "'"
        If SYMRead(gSL_Select, 1) Then
           mSTHopa = SYM.数字1 - 1
           mEDHopa = SYM.数字2 - 1
        Else
            Kn_Hopa = 5
        End If
    End If
    Pr_Cnt = ""

    DoEvents
    Call PrintStart
    
    Call Proc_End
        
    FormAct = Not FormAct
    
End Sub
'+----------------------------------------------------------+
'+      終了処理
'+----------------------------------------------------------+
Private Sub Proc_End()
    Unload CKN0240
End Sub
'+----------------------------------------------------------+
'+      発行種類確認
'+----------------------------------------------------------+
Private Function KanPrtSY(FirstFlg As Boolean) As Boolean
    KanPrtSY = False
    
    If Not FirstFlg And mKansy = KHT.発行種類 Then
       KanPrtSY = True
       Exit Function
    End If
    
    mKansy = KHT.発行種類
    kn_PrtSY = mKansy
    
    gSL_Select = "Select * From かんばん設定テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & mKansy
    If Not KSTRead(gSL_Select, 1) Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "      かんばんフォーマットが登録されていません。           "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        Call Proc_End
        Exit Function
    End If
    
    PrMsg(0) = "【  " & RTrim(KST.かんばん名) & "発行処理中 】"
    PrMsg(1) = PrMsg(0)

    ' レイアウト情報取得
    If Not FormatGet Then Exit Function
    
    KanPrtSY = True

End Function

'===========================================================================
'===========================================================================
'=======                                                             =======
'=======        【 かんばん発行処理開始 】                             =======
'=======                                                             =======
'===========================================================================
'===========================================================================
Private Sub PrintStart()
    Dim wPrtFlg     As Boolean
    
    With CisFrm
        Set .MeterPic = Bar
        .MeterMax = Kn_PCnt
        If Kn_Page Then
           If Kn_PageE = 0 Then
              Kn_PageE = Kn_PCnt
           End If
        End If
    End With

    ' RS232c 設定情報取得
    Call KnGetComm
    If Not Kn_TestFlg Then
        gLogFlag = kn_232cLog
        ' ログ出力開始
        Call LogStart
'   * Comm設定 *
        Call InitComm(MSComm1)
'   * 通信ポートの開局 *
        Call CommOpCl(MSComm1)
'
        gLong = 0
        Do Until MSComm1.CTSHolding
            gLong = gLong + 1
            If gLong > 4 Then
               With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "      処理を終了します。           "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
               End With
               Call Proc_End
               Exit Sub
            End If
            With CisFun
                 .MB_Lines = 5
                 .MB_MSG(2) = " 【かんばんプリンタの電源が投入されていません】     "
                 .MB_MSG(4) = "               電源を投入して下さい               "
                 .MB_Title = ""
                 .MB_Button = Retry
                 If Not .MBOX Then
                    Call Proc_End
                    Exit Sub
                 End If
            End With
        Loop
    End If
    ' かんばん発行処理
    gLong = 0
    Do Until Not KHT_RDSTS
    
        wPrtFlg = True
        
        If Kn_Page Then
           If gLong + 1 < Kn_PageS Then
              wPrtFlg = False
           End If
           If gLong + 1 > Kn_PageE Then Exit Do
        End If
        
        If wPrtFlg Then
        
            If Not KanPrtSY(False) Then Exit Sub
            DoEvents
            
            If kn_PrtKB <> 2 Then
                gSL_Select = "Select * From かんばん流動管理テーブル"
                gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHT.取引先) & "'"
                gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
                gSL_Select = gSL_Select & " And   品番 = '" & RTrim(KHT.品番) & "'"
                gSL_Select = gSL_Select & " And   発行NO = " & KHT.発行NO & ""
                If Not KRTRead(gSL_Select, 1) Then
                   Call ItemsClearKRT
                   With KRT
                        .取引先 = KHT.取引先
                        .受入 = KHT.受入
                        .品番 = KHT.品番
                        .発行NO = KHT.発行NO
                        .版数 = 1
                        .発行区分 = 0
                        .収容区分 = KHT.収容区分
                   End With
                   Call KRTInsert
                Else
                   With KRT
                        .版数 = .版数 + 1
                        .発行区分 = 0
                   End With
                   Call KRTUpdate
                End If
            Else
                gSL_Select = "Select * From かんばん材料管理テーブル"
                gSL_Select = gSL_Select & " Where 材料管理番号 = " & KHT.材料管理番号
                gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(KHT.取引先) & "'"
                gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
                gSL_Select = gSL_Select & " And   発行NO = " & KHT.発行NO & ""
                If Not KZKRead(gSL_Select, 1) Then
                   Call ItemsClearKZK
                   With KZK
                        .材料管理番号 = KHT.材料管理番号
                        .取引先 = KHT.取引先
                        .受入 = KHT.受入
                        .発行NO = KHT.発行NO
                        .版数 = 1
                        .発行区分 = 0
                        .収容区分 = KHT.収容区分
                   End With
                   Call KZKInsert
                Else
                   With KZK
                        .版数 = .版数 + 1
                        .発行区分 = 0
                   End With
                   Call KZKUpdate
                End If
            End If
            ' QRエリア転送
            Call QRAreaClear
            With QRArea
                .品番 = KHT.品番
                .品目 = KHT.品目
                .取引先 = KHT.取引先
                .受入 = KHT.受入
                .手配区分 = KHT.手配区分
                .版数 = KRT.版数
                .発行日 = Format(Now(), "YYYYMMDD")
                .背番号 = KHT.背番号
                .連番 = KHT.発行NO
                .収容数 = KHT.収容数
                .サイクル = KHT.サイクル
                .納入先 = KHT.納入先
                .納入先受入 = KHT.納入先受入
                .工程 = KHT.工程
                .次工程 = KHT.次工程
                .編集品番 = KHT.表示品番
                .手配区分名 = KHT.手配区分名
                .アドレス = KHT.アドレス
                If kn_PrtKB = 2 Then
                   .品目 = 2
                   .品番 = KHT.材質
                   .板厚 = KHT.板厚
                   .幅 = KHT.幅
                   .長さ = KHT.長さ
                   .寸法区分 = KHT.寸法区分
                   .材料管理番号 = KHT.材料管理番号
                   .材料収容数 = KHT.収容数
                End If
            End With
            '
            
            ' 発行
            If Kn_Hopa <> 5 Then
               KHT.StHNo = mSTHopa
               KHT.EdHNo = mEDHopa
               KHT.Stack1 = mSTHopa
               KHT.Stack2 = mSTHopa
            End If
            
            If Not FormatRSSet Then Exit Sub
            
            Call KanbanPrint
        
            With CisDB
                If kn_PrtKB <> 2 Then
                    .SQL = "Update かんばん発行テーブル Set 発行済 = 1,版数 = " & KRT.版数
                    .SQL = .SQL & " Where かんばん種類 = " & kn_PrtKB
                    .SQL = .SQL & " And   背番号 = '" & RTrim(KHT.背番号) & "'"
                    .SQL = .SQL & " And   品番 = '" & RTrim(KHT.品番) & "'"
                    .SQL = .SQL & " And   取引先 = '" & RTrim(KHT.取引先) & "'"
                    .SQL = .SQL & " And   受入 = '" & RTrim(KHT.受入) & "'"
                    .SQL = .SQL & " And   発行NO = " & KHT.発行NO
                Else
                    .SQL = "Update かんばん発行テーブル Set 発行済 = 1,版数 = " & KRT.版数
                    .SQL = .SQL & " Where かんばん種類 = " & kn_PrtKB
                    .SQL = .SQL & " And   材料管理番号 = " & KHT.材料管理番号
                    .SQL = .SQL & " And   背番号 = '" & RTrim(KHT.背番号) & "'"
                    .SQL = .SQL & " And   取引先 = '" & RTrim(KHT.取引先) & "'"
                    .SQL = .SQL & " And   受入 = '" & RTrim(KHT.受入) & "'"
                    .SQL = .SQL & " And   発行NO = " & KHT.発行NO
                End If
                .DBExec
            End With
        End If
        ' 進捗ゲージ更新
        gLong = gLong + 1
        CisFrm.MeterUpdate gLong
        Pr_Cnt = Format(gLong, "#,###") & " / " & Format(CisFrm.MeterMax, "#,###")
        
        DoEvents
        '中断処理
        If Prt_Cansel_Flg Then
            If MsgBox("処理を中断します。よろしいですか？", 33, "印刷") = 1 Then
                Exit Do
            Else
                Prt_Cansel_Flg = False
            End If
        End If
        
        Call KHTReadNext(1)
    Loop

    Call KHTClose(1)

'通信ポートの閉局 *
    If Not Kn_TestFlg Then
        Do
            DoEvents
            If Not XOnOffFlg Then
                Call CommOpCl(MSComm1, False)
                Exit Do
            End If
        Loop
        Call LogEnd
    End If
    
    Kn_Rtn = True
    
End Sub
'*******************************
'*      通信イベント管理        *
'*******************************
Private Sub MSComm1_OnComm()
    Static Rtn     As String
   
    Rtn = CommEvent(MSComm1)
'   * エラー *
    If Rtn = "ERR" Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(1) = gEventMsg
             .MB_MSG(2) = "      処理を終了します。           "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        End
        Exit Sub
    End If
    
    If kn_232cLog Then
       CisFun.LogPut (Rtn & ":" & gCommData)
    End If
    
    If Rtn = "RCV" Then
       For wXOnXOffCNT = 1 To Len(gCommData)
           If Mid(gCommData, wXOnXOffCNT, 1) = Chr(&H11) Then
              XOnOffFlg = False
              If kn_232cLog Then
                 CisFun.LogPut ("【X-On】")
              End If
           End If
           If Mid(gCommData, wXOnXOffCNT, 1) = Chr(&H13) Then
              XOnOffFlg = True
              If kn_232cLog Then
                 CisFun.LogPut ("【X-Off】")
              End If
           End If
       Next wXOnXOffCNT
    End If
End Sub
'+----------------------------+
'+    印刷取消キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Prt_Cansel_Flg = True
End Sub

