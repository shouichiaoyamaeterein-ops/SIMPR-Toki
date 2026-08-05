VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form CXH0333 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "発注データ e-Mail 送信処理"
   ClientHeight    =   1260
   ClientLeft      =   4365
   ClientTop       =   4050
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
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
   ScaleHeight     =   1260
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   570
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.PictureBox Bar 
      Height          =   240
      Left            =   105
      ScaleHeight     =   180
      ScaleWidth      =   5565
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   855
      Width           =   5625
   End
   Begin VB.PictureBox Picture1 
      Height          =   660
      Left            =   120
      ScaleHeight     =   600
      ScaleWidth      =   5535
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   5595
      Begin Cis3D_v60.CIS3D Proc_Msg 
         Height          =   555
         Left            =   45
         Top             =   30
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   979
         BackColor       =   12648384
         Caption         =   "e-Mail送信データ作成中"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   21.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
   End
End
Attribute VB_Name = "CXH0333"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注明細e-mail送信データ作成 & 送信
'**       フォームID    :   CXH0333
'**       処理概要      :
'**
'**       作  成  日    :   2004/05/14  By CIS
'**       変  更  日    :   2008/04/11  By CIS - 背番号８桁対応
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替

    Dim mShiire                 As String
    
    Dim mPath                   As String
    Dim mFName                  As String
    Dim mFKotei                 As String
    Dim mFNo                    As Integer

    Dim mMsgPath                As String
    Dim mMsgFileBase            As String
    Dim mMsgFileName            As String
    
    Dim mTextRec                As String
    
    Dim mTorcd()                As String
    Dim mJN_Subject             As String
    Dim mJN_MSGText             As String
    Dim FNo                     As Byte
    Dim PosCnt                  As Long
    Dim mJN_MailChk             As Boolean
    
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》            +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    
    Me.MousePointer = 11
    
    Proc_Msg = "e-Mail送信データ作成中"
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(False, x1024y768)

End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct
    
    Call ProcStart
    
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
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
'    Key_Acc = False
'
'    Select Case Key_Code
'           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
'           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click        '【終了】
'           Case vbKeyReturn
'                                If PB_ENT.Visible Then                          '【入力】
'                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
'                                    Call PB_ENT_Click
'                                    Dummy.Enabled = False
'                                End If
'           Case Else:           Exit Function
'    End Select
'
'    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Me.Hide
End Sub
'+---------------------+
'+      処    理       +
'+---------------------+
Private Sub ProcStart()
    '
    ' 送信データ作成
    '
    With CisFrm
        Set .MeterPic = Bar
        .MeterMax = CXH0330_PrtCnt
    End With
    
    Call FileInit
    ReDim mTorcd(1)
    mTorcd(0) = HCT.仕入先
    mShiire = HCT.仕入先
    If Not OutFileOpen Then GoTo ProcStart_ed
    If Not HeaderPut Then GoTo ProcStart_ed
    
    gCnt1 = 0
    Do Until Not HCT_RDSTS
        gCnt1 = gCnt1 + 1
        If mShiire <> HCT.仕入先 Then
            If Not TrailerPut Then GoTo ProcStart_ed
            gCnt1 = 1
            mShiire = HCT.仕入先
            Close #mFNo
                            
            gCnt2 = UBound(mTorcd)
            ReDim Preserve mTorcd(gCnt2 + 1)
            mTorcd(gCnt2) = HCT.仕入先
            
            If Not OutFileOpen Then GoTo ProcStart_ed
            If Not HeaderPut Then GoTo ProcStart_ed
        End If
    '------------------------'
    '   レコード編集
    '------------------------'
        Call HCTextEdit
    '------------------------'
    '   レコード出力
    '------------------------'
'        On Error GoTo TextFile_CrtProc_Err
        Print #mFNo, mTextRec
    
        CisFrm.MeterUpdate gCnt1
    
        Call HCTReadNext(1)
    Loop
    
    CisFrm.MeterUpdate CisFrm.MeterMax
    
    If Not TrailerPut Then GoTo ProcStart_ed
    
    Close #mFNo
    
    Call HCTClose(1)

    '
    ' Mail送信
    '
    For gCnt1 = 0 To UBound(mTorcd)
        If Trim(mTorcd(gCnt1)) = "" Then Exit For
        If Not MailSendProc(mTorcd(gCnt1)) Then GoTo ProcStart_ed
    Next gCnt1


ProcStart_ed:
    ' 処理終了
    Call PB_END_Click
End Sub
'+------------------------------+
'+      添付ファイル処理        +
'+------------------------------+
'+-----------------------------------------------+
'+      初期処理(ﾌｫﾙﾀﾞ確認 & 前回処理ﾌｧｲﾙ削除)
'+-----------------------------------------------+
Private Sub FileInit()
    
    mFKotei = "HACHU"
    
    '------------------------'
    '   ディレクトリ存在確認
    '------------------------'
    With CisFun
        .INI_Section = "Mail"
        .INI_FileID = gIniExe & gIniName
        ' ﾒｯｾｰｼﾞﾌｧｲﾙ格納先
        .INI_Default = ""
        .INI_Key = "MailText"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = gIniExe & "Mail\"
           .PutIni
        End If
        mPath = Trim(.INI_String)
        If Right(mPath, 1) <> "\" Then mPath = mPath & "\"
        
    End With
    With CisFun
        .FE_Path = mPath
        .DirCheck NoMsg_Crt
    End With
    '------------------------'
    '   前回処理ﾌｧｲﾙ削除
    '------------------------'
    On Error Resume Next
    mFName = mPath & mFKotei & "*" & ".txt"
    Kill mFName
    On Error GoTo 0
End Sub
'+-----------------------------+
'+    出力ファイルオープン
'+-----------------------------+
Private Function OutFileOpen() As Boolean
    OutFileOpen = False
    
    On Error GoTo OutFileOpen_er:
    
    If mFNo <> 0 Then
       Close #mFNo
    End If

    mFNo = FreeFile
    mFName = mPath & mFKotei & mShiire & ".txt"
    Open mFName For Output As #mFNo
    
    
    OutFileOpen = True
    Exit Function
OutFileOpen_er:
    With CisFun
        .MB_Lines = 5
        .MB_MSG(1) = "      出力ファイルオープンエラー      "
        .MB_MSG(3) = "   ファイル = " & mFName
        .MB_MSG(5) = " Err=" & Err & ",Error=" & Error
        .MB_Title = "ｵｰﾌﾟﾝｴﾗｰ"
        .MB_Button = OK
        .MBOX
    End With
End Function
'+-----------------------------+
'+    ヘッダー出力
'+-----------------------------+
Private Function HeaderPut() As Boolean
    HeaderPut = False
    On Error GoTo HeaderPut_Er
        
    If mFNo = 0 Then Exit Function
    
    mTextRec = "H" & Format(Date, "YYYYMMDD") & mShiire
    Print #mFNo, mTextRec
    
    HeaderPut = True
    Exit Function
HeaderPut_Er:
    With CisFun
        .MB_Lines = 5
        .MB_MSG(1) = "      ヘッダー出力エラー      "
        .MB_MSG(3) = "   ファイル = " & mFName
        .MB_MSG(5) = " Err=" & Err & ",Error=" & Error
        .MB_Title = "出力ｴﾗｰ"
        .MB_Button = OK
        .MBOX
    End With
End Function
'+-----------------------------+
'+    トレーラー出力
'+-----------------------------+
Private Function TrailerPut() As Boolean
    TrailerPut = False
    On Error GoTo TrailerPut_Er
        
    If mFNo = 0 Then Exit Function
    
    mTextRec = "T" & Format(gCnt1, "0000")
    Print #mFNo, mTextRec
    
    TrailerPut = True
    Exit Function
TrailerPut_Er:
    With CisFun
        .MB_Lines = 5
        .MB_MSG(1) = "      トレーラー出力エラー      "
        .MB_MSG(3) = "   ファイル = " & mFName
        .MB_MSG(5) = " Err=" & Err & ",Error=" & Error
        .MB_Title = "出力ｴﾗｰ"
        .MB_Button = OK
        .MBOX
    End With
End Function

'************************************
'*    発注テキスト編集
'************************************
Private Function HCTextEdit() As Boolean
    HCTextEdit = False

    mTextRec = Space(250)
    Mid(mTextRec, 1, 1) = "D"                              'データ区分
    With HCT
        Mid(mTextRec, 2, 5) = RTrim(.納入番号)                             '納入番号
        Mid(mTextRec, 7, 2) = Format(.納品書頁, "00")                      '納品書頁
        Mid(mTextRec, 9, 3) = Format(.納品書行, "000")                     '納品書行
        Mid(mTextRec, 12, 1) = Format(.分納区分, "0")                      '分納区分
        Mid(mTextRec, 13, 7) = RTrim(.仕入先)                              '手配先
        Mid(mTextRec, 20, 7) = RTrim(.契約先)                              '契約先
        Mid(mTextRec, 27, 2) = RTrim(.受入)                                '受入
        If RTrim(.表示品番) <> "" Then
            Mid(mTextRec, 29, 25) = CisFun.Left2(.表示品番, 25)                '品番
        Else
            Mid(mTextRec, 29, 25) = CisFun.Left2(.品番, 25)                '品番
        End If
        gStr = Space(6)
'        If Not IsNumeric(Trim(gStr)) Then
'            gStr = "000000"
'        End If
        Mid(mTextRec, 54, 6) = CisFun.Left2(gStr, 6)                       '空き

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
'        Mid(mTextRec, 60, 5) = RTrim(.背番号)                               '背番号
'        Mid(mTextRec, 65, 8) = RTrim(.納入日)                               '納入日
'        Mid(mTextRec, 73, 2) = RTrim(.納入便)                               '納入便
'        Mid(mTextRec, 75, 6) = Format(.収容数, "000000")                    '収容数
'        Mid(mTextRec, 81, 3) = Format(.枚数, "000")                         '枚数
'        Mid(mTextRec, 84, 6) = Format(.端数, "000000")                      '端数
'        Mid(mTextRec, 90, 6) = Format(.発注数, "000000")                    '発注数
'        Mid(mTextRec, 96, 1) = Format(.発注区分, "0")                      '発注区分
'        Mid(mTextRec, 97, 1) = Format(.手配区分, "0")                      '手配区分
'        Mid(mTextRec, 98, 1) = Format(.SK区分, "0")                        'SK区分
'        Mid(mTextRec, 99, 7) = RTrim(.納入先)                              '納入先
'        Mid(mTextRec, 106, 2) = RTrim(.納入先受入)                          '納入先受入
'        Mid(mTextRec, 108, 1) = RTrim(.作成種別)                            '処理区分(A:追加、M:修正、D:削除)
'        Mid(mTextRec, 109, 12) = RTrim(.作成日付)                           '作成日付
'        Mid(mTextRec, 121, 12) = RTrim(.更新日付)                          '更新日付
        Mid(mTextRec, 60, 8) = RTrim(.背番号)                               '背番号
        Mid(mTextRec, 68, 8) = RTrim(.納入日)                               '納入日
        Mid(mTextRec, 76, 2) = RTrim(.納入便)                               '納入便
        Mid(mTextRec, 78, 6) = Format(.収容数, "000000")                    '収容数
        Mid(mTextRec, 84, 3) = Format(.枚数, "000")                         '枚数
        Mid(mTextRec, 87, 6) = Format(.端数, "000000")                      '端数
        Mid(mTextRec, 93, 6) = Format(.発注数, "000000")                    '発注数
        Mid(mTextRec, 99, 1) = Format(.発注区分, "0")                      '発注区分
        Mid(mTextRec, 100, 1) = Format(.手配区分, "0")                      '手配区分
        Mid(mTextRec, 101, 1) = Format(.SK区分, "0")                        'SK区分
        Mid(mTextRec, 102, 7) = RTrim(.納入先)                              '納入先
        Mid(mTextRec, 109, 2) = RTrim(.納入先受入)                          '納入先受入
        Mid(mTextRec, 111, 1) = RTrim(.作成種別)                            '処理区分(A:追加、M:修正、D:削除)
        Mid(mTextRec, 112, 12) = RTrim(.作成日付)                           '作成日付
        Mid(mTextRec, 124, 12) = RTrim(.更新日付)                          '更新日付
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    End With

    HCTextEdit = True
End Function
'+-----------------------------------------------+
'+      Mail送信処理
'+-----------------------------------------------+
Private Function MailSendProc(Tor As String) As Boolean
    MailSendProc = False

    Dim Send_Torcd(4)       As String
    Dim Send_Name(4)        As String
    Dim Send_Cmpany(4)      As String
    Dim Send_Busyo(4)       As String
    Dim Send_Mail(4)        As String
    Dim Send_Msg(4)         As String

    '送信先情報を設定
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(Tor) & "'"
    If TRMRead(gSL_Select, 1, 1) Then
       If Trim(TRM.EMAIL) <> "" Then
          Send_Torcd(0) = Tor
          Send_Name(0) = TRM.担当者
          Send_Cmpany(0) = TRM.略称
          Send_Busyo(0) = TRM.部署名1
          Send_Mail(0) = TRM.EMAIL
          Send_Msg(0) = TRM.MSGTYPE
          If Trim(Send_Name(0)) = "" Then
             Send_Name(0) = Send_Cmpany(0)
          End If
       End If
       Send_Torcd(1) = TRM.MAIL取引先1
       Send_Torcd(2) = TRM.MAIL取引先2
       Send_Torcd(3) = TRM.MAIL取引先3
    End If
    
    For gCnt2 = 1 To 3
        If Trim(Send_Torcd(gCnt2)) = "" Then Exit For
        gSL_Select = "SELECT * FROM 取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(Send_Torcd(gCnt2)) & "'"
        If TRMRead(gSL_Select, 1, 1) Then
           If Trim(TRM.EMAIL) <> "" Then
              Send_Torcd(gCnt2) = Send_Torcd(gCnt2)
              Send_Name(gCnt2) = TRM.担当者
              Send_Cmpany(gCnt2) = TRM.略称
              Send_Busyo(gCnt2) = TRM.部署名1
              Send_Mail(gCnt2) = TRM.EMAIL
              Send_Msg(gCnt2) = TRM.MSGTYPE
              If Trim(Send_Name(gCnt2)) = "" Then
                 Send_Name(gCnt2) = Send_Cmpany(gCnt2)
              End If
           End If
        End If
    Next gCnt2
    
    If Trim(Send_Torcd(0)) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "    送信先E-mailが登録されていません     "
            .MB_MSG(4) = "      送信先 = " & Tor
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

    '------------------------'
    '   ディレクトリ存在確認
    '------------------------'
    ' Mail送信情報
    With CisFun
        .INI_Section = "Mail"
        .INI_FileID = gIniExe & gIniName
        ' ﾒｯｾｰｼﾞﾌｧｲﾙ格納先
        .INI_Default = ""
        .INI_Key = "MsgDir"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = gIniExe & "Mail\"
           .PutIni
        End If
        mMsgPath = Trim(.INI_String)
        If Right(mMsgPath, 1) <> "\" Then mMsgPath = mMsgPath & "\"
        
        .INI_Default = ""
        .INI_Key = "Mail即時送信"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Y"
           .PutIni
        End If
        If UCase(.INI_String) = "Y" Or UCase(.INI_String) = "YES" Then
           mJN_MailChk = False
        Else
           mJN_MailChk = True
        End If
    
    End With
    
    With CisFun
        .FE_Path = mMsgPath
        .DirCheck NoMsg_Crt
    End With
    mMsgFileBase = "MailMsg"
    
    For gCnt2 = 0 To 3
        If Trim(Send_Torcd(gCnt2)) = "" Then Exit For
        
        mJN_Subject = ""
        mJN_MSGText = ""
        
        mMsgFileName = mMsgPath & mMsgFileBase & Trim(Send_Msg(gCnt2)) & ".txt"
        If RTrim$(Dir(mMsgFileName)) = "" Then
           mMsgFileName = mMsgPath & mMsgFileBase & ".txt"
        Else
            If FileLen(mMsgFileName) = 0 Then
               mMsgFileName = mMsgPath & mMsgFileBase & ".txt"
            End If
        End If
        If RTrim$(Dir(mMsgFileName)) <> "" Then
            If FileLen(mMsgFileName) <> 0 Then
                FNo = FreeFile
                
                Open mMsgFileName For Input Access Read As #FNo
                
                Do
                    Line Input #FNo, gStr
                    gStr = Replace(gStr, Chr(9), " ")   ' TAB消去
                    
                    If Mid(gStr, 1, 2) <> "--" Then
                        
                        If RTrim$(gStr) <> "" Then
                            If InStr(gStr, "<YYYY>") <> 0 Then gStr = Replace(gStr, "<YYYY>", Format(Now(), "yyyy"))
                            If InStr(gStr, "<YY>") <> 0 Then gStr = Replace(gStr, "<YY>", Format(Now(), "yy"))
                            If InStr(gStr, "<MM>") <> 0 Then gStr = Replace(gStr, "<MM>", Format(Now(), "mm"))
                            If InStr(gStr, "<DD>") <> 0 Then gStr = Replace(gStr, "<DD>", Format(Now(), "dd"))
                            If InStr(gStr, "<送信先会社名>") <> 0 Then gStr = Replace(gStr, "<送信先会社名>", Send_Cmpany(gCnt2))
                            If InStr(gStr, "<送信先担当者>") <> 0 Then gStr = Replace(gStr, "<送信先担当者>", Send_Name(gCnt2))
                            If InStr(gStr, "<送信先部署>") <> 0 Then gStr = Replace(gStr, "<送信先部署>", Send_Busyo(gCnt2))
                            If InStr(gStr, "<仕入先会社名>") <> 0 Then gStr = Replace(gStr, "<仕入先会社名>", Send_Cmpany(0))
                            If InStr(gStr, "<仕入先CD>") <> 0 Then gStr = Replace(gStr, "<仕入先CD>", Send_Torcd(0))
        '                    If InStr(gStr, "<再送>") <> 0 Then
        '                       If SJR.送信回数 > 0 Then
        '                          gStr = Replace(gStr, "<再送>", "[再送]")
        '                       Else
        '                          gStr = Replace(gStr, "<再送>", "")
        '                       End If
        '                    End If
                        End If
                        If Mid(gStr, 1, 2) = "件名" Then
                           PosCnt = InStr(1, gStr, "：")
                           If PosCnt > 0 Then
                               mJN_Subject = Mid(gStr, PosCnt + 1)
                           End If
                        Else
                            mJN_MSGText = mJN_MSGText & gStr & vbCrLf
                        End If
                    End If
                Loop Until EOF(FNo)
                
                Close #FNo
            End If
        End If
    '   < メッセージ作成用ワークの初期化 >
        Call gMapi_ItemClear
    
        gMapi_MailTO = Send_Mail(gCnt2)      ' 送信先(宛先)
        gMapi_MailCC = ""               ' CC(カーボンコピー)
        gMapi_MailSBJ = mJN_Subject     ' 件名
        gMapi_MailMSG = mJN_MSGText     ' 本文
        
        gMapi_AttcBIN(0) = mPath & mFKotei & Trim(Tor) & ".txt"
    ' 添付ファイル
        If RTrim$(Dir(gMapi_AttcBIN(0))) = "" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    添付ファイルが存在しません。     "
                .MB_MSG(4) = "      ファイル名 = " & gMapi_AttcBIN(0)
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            Exit Function
        End If
    
        If Not SndMail(MAPISession1, MAPIMessages1, mJN_MailChk) Then Exit Function
    Next gCnt2


    MailSendProc = True
    Exit Function
MailSendProc_ER:

End Function
