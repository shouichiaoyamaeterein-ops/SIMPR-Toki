VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Begin VB.Form CCT0030 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "SIMPR-A"
   ClientHeight    =   6900
   ClientLeft      =   2670
   ClientTop       =   2010
   ClientWidth     =   10110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   6900
   ScaleWidth      =   10110
   Begin Cis3D_v60.CIS3D B1_TextProc 
      Height          =   495
      Left            =   210
      Top             =   720
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   873
      BackColor       =   0
      ForeColor       =   65535
      Caption         =   "受 信 処 理 開 始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cFont3D         =   2
      cAlingnment     =   7
   End
   Begin VB.PictureBox B1_Gage 
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   210
      ScaleHeight     =   225
      ScaleWidth      =   9615
      TabIndex        =   2
      Top             =   1290
      Width           =   9675
   End
   Begin Cis3D_v60.CIS3D B1_TextFMEI 
      Height          =   525
      Left            =   210
      Top             =   90
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   926
      Caption         =   "【 】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   8421504
      cFont3DColor2   =   -2147483643
      cBoderWidth     =   3
      cFont3D         =   1
      cAlingnment     =   7
   End
   Begin VB.TextBox Dummy 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1890
      Width           =   195
   End
   Begin VB.ListBox ProcList 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5010
      ItemData        =   "CCT0030.frx":0000
      Left            =   210
      List            =   "CCT0030.frx":0002
      TabIndex        =   0
      Top             =   1680
      Width           =   9645
   End
End
Attribute VB_Name = "CCT0030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  テキスト変換
'**       フォームID    :  CCT0030
'**       処理概要      :
'**
'**       作  成  日    :  2003/12/10  By CIS
'**       変  更  日    :  2004/11/09  BY CIS 財務ﾃﾞｰﾀ取込処理追加
'**       変  更  日    :  2005/03/31  By CIS 計算方法を取引先マスタの設定に従う
'**       変  更  日    :  2007/03/05  By CIS 手配ﾃﾞｰﾀ取込処理追加
'**       変  更  日    :  2008/05/09  By CIS 車体精工対応
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
 
    Dim cIDX            As Long
    Dim cIDX2           As Long
    
    Dim mRirekiNo       As Long         ' 履歴管理番号
    
    Dim mFirstHead      As String       ' ﾍｯﾀﾞｰ内容（最初のﾍｯﾀﾞｰ)
    Dim mChackHead      As String       ' ﾍｯﾀﾞｰ内容（ﾁｪｯｸ用のﾍｯﾀﾞｰ)
'
    Dim mFNo            As Integer      ' 入力ﾌｧｲﾙ№
    Dim mInFName        As String       ' 入力ﾌｧｲﾙ名

    Dim mFileByte       As Long         ' ﾌｧｲﾙﾊﾞｲﾄ数
'
    Dim mErrStr1        As String
    Dim mErrStr2        As String
    Dim mErr            As String
'
    Dim mHeadChkFlg     As Boolean      ' ﾍｯﾀﾞｰ処理済確認ﾌﾗｸﾞ
    Dim mCompany        As String       ' 自社コード
    Dim mMonFlg         As Long

    Dim mCnt1           As Long
    
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
    B1_TextFMEI = ""
    ProcList.Clear
    Set ProcListObj = ProcList
'   #---------------------------------#
'   #  テキスト変換処理初期情報　取得   #
'   #---------------------------------#
    Call TextInit
    
    ReDim typText.ERRCODE(UBound(typText.FILECD))
    ReDim typText.HFName(UBound(typText.FILECD))
'   内示工区変換情報取得
    ' 工区変換
    Erase tx_NjKouku
    Erase tx_NjKoukuX
    gSL_Select = "Select 値,英数字3 from 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '内示工区変換' "
    gSL_Select = gSL_Select & "   AND 区分タイプ = '' "
    gSL_Select = gSL_Select & "   AND 値 <> '' "
    If SYMRead(gSL_Select) Then
       ReDim tx_NjKouku(CisDB.RecordCount)
       ReDim tx_NjKoukuX(CisDB.RecordCount)
       txCnt1 = 0
       Do Until Not SYM_RDSTS
          tx_NjKouku(txCnt1) = Mid(SYM.値, 1, 8)
          tx_NjKoukuX(txCnt1) = Mid(SYM.英数字3, 1, 1)
              
          txCnt1 = txCnt1 + 1
          Call SYMReadNext
       Loop
    Else
       ReDim tx_NjKouku(1)
       ReDim tx_NjKoukuX(1)
    End If
    Call SYMClose

    Dummy.Left = -1000
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CISFrm.Frm_Center(False, x1024y768)
    
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》           +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = True
    
    DoEvents
    
    Dummy.SetFocus
    
    Call ReturnPress
    
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＰｒｅｓｓイベント 》             +
'+----------------------------------------------------------+
Private Sub Form_KeyPress(KeyAscii As Integer)
'   *  入力項目値チェック   *
    If Not (ActiveControl.Name Like "i*") Then Exit Sub

End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
End Function
'+--------------------------+
'+      終  了 処  理       +
'+--------------------------+
Private Sub EndProc(Optional EndType As String)
    Me.Hide
End Sub
'**********************************
'*     テキスト変換・累積処理       *
'**********************************
Private Sub ReturnPress()
    Dim wDelDate        As String
    Me.MousePointer = 11
    

    Call CisDB.DBTran(TransBegin, 0)
    Call CisDB.DBTran(TransBegin, 1)
'-------------------------
'-    履歴テーブル削除
'-------------------------
    On Error Resume Next
    For cIDX = 0 To UBound(typText.FILECD) - 1
        With CisDB
            .SQL = "DELETE  TEXT変換履歴テーブル "
            .SQL = .SQL & " Where FILECD = '" & RTrim(typText.FILECD(cIDX)) & "'"
            .SQL = .SQL & "   And 作成日 <= '" & Format(Now() - typText.LogDelDay(cIDX), "YYYY/MM/DD hh:nn:ss") & "'"
'            .SQL = .SQL & "   And 作成日 <= '" & Format(Now() - tx_LogDelDay, "YYYY/MM/DD hh:nn:ss") & "'"
            
            Call .DBExec
        End With
    Next cIDX
    On Error GoTo ReturnPress_Ed
'-------------------------
'-    変 換 処 理
'-------------------------
    For cIDX = 0 To UBound(typText.FILECD) - 1
       If Not ProcHenkan(cIDX) Then
          Select Case typText.ERRCODE(cIDX)
                Case "H6": gStr = "中止"
                Case Else: gStr = "ｴﾗｰ"
          End Select
          tx_GridObj.TextMatrix(typText.ProcRow(cIDX), tx_ErMsg) = gStr
          tx_GridObj.TextMatrix(typText.ProcRow(cIDX), tx_Result) = typText.ERRCODE(cIDX)
          tx_ProcErr = True
       Else
          tx_GridObj.TextMatrix(typText.ProcRow(cIDX), tx_ErMsg) = ""
       End If
    ' Text管理テーブル更新
       Call TEXTUpdate(typText.ERRCODE(cIDX), cIDX)
    ' 履歴情報出力
       Call TextRirekiPut(mRirekiNo, "E", typText.ERRCODE(cIDX))
    Next cIDX
    
ReturnPress_Ed:
    Me.MousePointer = 1
    If tx_ProcErr Then
        Call CisDB.DBTran(TransRollback, 0)
    Else
        Call CisDB.DBTran(TransCommit, 0)
    End If
    Call CisDB.DBTran(TransCommit, 1)
    
    Call EndProc
    
End Sub
'+-------------------------------+
'+-------------------------------+
'+-------------------------------+
'+       変 換 処 理              +
'+-------------------------------+
'+-------------------------------+
'+-------------------------------+
Private Function ProcHenkan(HIdx As Long) As Boolean
    Dim FSeek           As Long
    Dim ByteArray()     As Byte
    Dim DataArray()     As String
    Dim GetID           As Long
    Dim ErrType         As String
    Dim FLeng           As Long
    
    ProcHenkan = False
    
    On Error GoTo ProcHenkan_Er
    '+------------ 変換情報初期化
    tx_SyoriCnt = 0
    nj_YYMM = ""
    nj_Renban = 0
    
    Call ItemsClearTKT
    TKT.FILECD = typText.FILECD(HIdx)
    ' 履歴情報出力
    Call TextRirekiPut(mRirekiNo, "S", "")
    
    If Not TKTRead("", 1) Then
       Call ProcLogPut("ファイルコード未設定(変換処理)")
       typText.ERRCODE(HIdx) = "H3"
       Call TERPut("H3")
       Exit Function            ' Textﾌｧｲﾙ未設定
    End If
    
                            ' 変換元フォルダ
    With CisFun
        .INI_FileID = gIniExe & gwsIniName
        .INI_Section = "TextConvert"
        .INI_Default = gIniExe & "Text\"
        .INI_Key = "InputDir_" & typText.FILECD(HIdx)
        .GetIni
        If Trim(.INI_String) <> "" Then
           tx_InputDir = RTrim(.INI_String)
           If Right(tx_InputDir, 1) <> "\" Then tx_InputDir = tx_InputDir & "\"
'           ' 自社情報確認チェック
'           .INI_Section = "TextConvert"
'           .INI_Default = "Yes"
'           .INI_Key = "Jisya_" & typText.FILECD(HIdx)
'           .GetIni
'           If UCase(.INI_String) = "Y" Or UCase(.INI_String) = "YES" Then
'              tx_JisyaChk = True
'           Else
'              tx_JisyaChk = False
'           End If
        End If
    
        ' 自社情報確認チェック
        .INI_Section = "TextConvert"
        .INI_Default = "Yes"
        .INI_Key = "Jisya_" & typText.FILECD(HIdx)
        .GetIni
        If UCase(.INI_String) = "Y" Or UCase(.INI_String) = "YES" Then
           tx_JisyaChk = True
        Else
           tx_JisyaChk = False
        End If
    End With
    
    If TKT.TNSタイプ = 0 Then
        ' レイアウト情報を求める
        gSL_Select = "Select 数字1 From 名称マスタ "
        gSL_Select = gSL_Select & " Where 区分名称 = 'レイアウト情報内示'"
        gSL_Select = gSL_Select & "   And 値 = '" & TKT.TEXTレイアウト & "'"
        If Not SYMRead(gSL_Select, 1) Then
           Call ItemsClearSYM
        End If
        tx_LayoutType = SYM.数字1
' テキスト設定情報　取得
        If Not typConvertSet(mErr) Then
           typText.ERRCODE(HIdx) = mErr
           Exit Function            ' Textﾌｧｲﾙ未設定
        End If
    End If
' Text変換元ﾌｧｲﾙ ファイル名を求める
    If Trim(typText.HFName(HIdx)) = "" Then
        typText.HFName(HIdx) = TextMaxFName(TKT.変換元ファイル名)
    End If
       
    typText.HFName(HIdx) = tx_InputDir & RTrim(typText.HFName(HIdx))
    mInFName = RTrim(typText.HFName(HIdx))
    
    tx_InCnt = 0: tx_BunCnt = 0
    mFileByte = 0: mFirstHead = ""
    '
    B1_TextFMEI = "【 " & Trim(TKT.F名称) & " 】"
    B1_TextProc = "変 換 処 理 開 始"
    Call ProcLogPut("変換処理開始 変換元ﾌｧｲﾙ名 = " & mInFName)
    '
    If Trim(mInFName) = "" Then
       Call ProcLogPut("変換元テキストファイル未登録(変換処理)")
       typText.ERRCODE(HIdx) = "H2"
       Call TERPut("H2")
       GoTo ProcHenkan_E1
    End If
    
    On Error Resume Next
    ' 変換元ファイル存在チェック
    If Dir(mInFName) = "" Then
       Call ProcLogPut("テキストデータファイル異常。(ﾌｧｲﾙ未登録)　＜変換処理＞ ﾌｧｲﾙ名=" & mInFName)
       mErrStr1 = "     テキストデータファイルが異常です     "
       mErrStr2 = "           (ファイル未登録)            " & vbCrLf & vbCrLf & _
                  "         ﾌｧｲﾙ名=" & mInFName
           
       typText.ERRCODE(HIdx) = "H0"
       Call TERPut("H0", mErrStr1, mErrStr2)
       GoTo ProcHenkan_E1
    End If
            
    mFileByte = FileLen(mInFName)
    If mFileByte = 0 Then
       Call ProcLogPut("テキストデータファイル異常。(長さ０バイト)　＜変換処理＞")
       
       typText.ERRCODE(HIdx) = "H1"
       Call TERPut("H1")
       GoTo ProcHenkan_E1
    End If
    
    On Error GoTo 0
    ' 非自社ｺｰﾄﾞ取得
    If tx_JisyaChk Then
       '--------------------------------------------- 2007/03/12 Delete
'       gSL_Select = "Select 被自社CD From 取引先マスタ "
'       gSL_Select = gSL_Select & " Where 取引先CD = '" & TKT.契約先 & "'"
'       gSL_Select = gSL_Select & "   And 取引先区分 = 0 "
'       gSL_Select = gSL_Select & "   And 契約先区分 = 1 "
'       If TRMRead(gSL_Select, 1) Then
'          mCompany = Trim(TRM.被自社CD)
'       End If
       '--------------------------------------------- 2007/03/12 Delete End
       '--------------------------------------------- 2007/03/12 Insert
       mCompany = RTrim(gCompany)
       gSL_Select = "Select 被自社CD From 取引先マスタ "
       gSL_Select = gSL_Select & " Where 取引先CD = '" & TKT.契約先 & "'"
       If TKT.TNSタイプ <> 2 Then
          gSL_Select = gSL_Select & "   And 取引先区分 = 0 "
       Else
          gSL_Select = gSL_Select & "   And 取引先区分 = 1 "
       End If
       gSL_Select = gSL_Select & "   And 契約先区分 = 1 "
       If TRMRead(gSL_Select, 1) Then
          If Trim(TRM.被自社CD) <> "" Then
             mCompany = Trim(TRM.被自社CD)
          End If
       End If
       '--------------------------------------------- 2007/03/12 Insert End
    End If
    '------------------------------------------------- 2007/03/12 Insert
    If Trim(TKT.自社CD) <> "" Then
       mCompany = TKT.自社CD
    End If
    '------------------------------------------------- 2007/03/12 Insert End
    '+-----------------------------------------------  2005/03/31 Insert
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & TKT.契約先 & "'"
    '-----------------------------------------------   2006/03/12 Insert
    If TKT.TNSタイプ <> 2 Then
       gSL_Select = gSL_Select & "   And 取引先区分 = 0 "
    Else
       gSL_Select = gSL_Select & "   And 取引先区分 = 1 "
    End If
    '-----------------------------------------------   2006/03/12 Insert End
    '-----------------------------------------------   2006/03/12 Delete
'    gSL_Select = gSL_Select & "   And 取引先区分 = 0 "
    '-----------------------------------------------   2006/03/12 Delete End
    gSL_Select = gSL_Select & "   And 契約先区分 = 1 "
    If Not TRMRead(gSL_Select, 1) Then
       Call ItemsClearTRM
    End If
    '+-----------------------------------------------  2005/03/31 Insert End
    
    '
    B1_TextProc = "変 換 処 理 中"
    DoEvents
    
    tx_InCnt = 0: tx_HeadCnt = 0: tx_TailCnt = 0: tx_BunCnt = 0: tx_OutCnt = 0
    tx_KNo = 0
    Erase tx_Shiki: ReDim tx_Shiki(1)
    
    If TKT.TNSFMT = 0 Then
       FLeng = 1000
    Else
       FLeng = TKT.有効長
    End If
    If FLeng = 0 Then
       FLeng = 250
    End If
    
    
    mHeadChkFlg = False  ' ﾍｯﾀﾞｰ処理済確認ﾌﾗｸﾞ初期化
    
    On Error GoTo ProcHenkan_Er
    mFNo = FreeFile
    
    If TKT.オプション1 = 0 Then     ' 行区切り無し
       With CISFrm
            Set .MeterPic = B1_Gage
            .MeterMax = mFileByte
       End With
       
       Open mInFName For Binary Access Read As #mFNo
        
       ReDim ByteArray(FLeng - 1)
       FSeek = 1
       Do
          Get #mFNo, FSeek, ByteArray()
          ReDim DataArray(FLeng - 1)
           
          tx_DataStr = ""
          tx_DataStr = StrConv(ByteArray(), vbUnicode)
          
          tx_DataStr = Replace(tx_DataStr, Chr(0), " ") ' Null文字削除
          
          If Not TextCheck(HIdx) Then GoTo ProcHenkan_E1
           
          
          FSeek = Seek(mFNo)
          ' 進捗ゲージ更新
          CISFrm.MeterUpdate FSeek
          DoEvents
        
        Loop While FSeek < mFileByte
    Else                            ' 行区切りあり
       With CisFun
            .FE_FullPath = mInFName
            mFileByte = .FileLine
       End With
'       If mFileByte < 3 Then           '2010/02/15
       If mFileByte < 1 Then            '2010/02/15
          Call ProcLogPut("テキストデータファイル異常。(改行がありません)　＜変換処理＞")
              
          typText.ERRCODE(HIdx) = "H5"
          Call TERPut("H5")
          GoTo ProcHenkan_E1
       End If
        
       With CISFrm
            Set .MeterPic = B1_Gage
            .MeterMax = mFileByte
       End With
       Open mInFName For Input As #mFNo
       Do While Not EOF(mFNo)
          Line Input #mFNo, tx_DataStr
        
          If Not TextCheck(HIdx) Then GoTo ProcHenkan_E1
          
          ' 進捗ゲージ更新
          FSeek = tx_SyoriCnt
'          If FSeek > mFileByte Then FSeek = mFileByte
          CISFrm.MeterUpdate FSeek
          DoEvents
       Loop
    End If
    ' 進捗ゲージ更新(最大値)
    CISFrm.MeterUpdate mFileByte
    '
    Close #mFNo
'
    If TKT.TNSタイプ = 0 And Trim(nj_YYMM) <> "" Then
        ' 内示対象年月更新
        gSL_Select = "Select 英数字1 From 名称マスタ "
        gSL_Select = gSL_Select & " Where 区分名称 = '内示年月'"
        gSL_Select = gSL_Select & " And   区分タイプ = ''"
        gSL_Select = gSL_Select & " And   値 = '1'"
        If SYMRead(gSL_Select, 1) Then
            With CisDB
                .SQL = "Update 名称マスタ Set "
                .SQL = .SQL & " 英数字1 = '" & nj_YYMM & "',"
                .SQL = .SQL & " 数字1 = '" & TKT.オプション3 & "',"
                .SQL = .SQL & "更新日 = GETDATE(),"
                .SQL = .SQL & "更新者 = '" & gTanto & "',"
                .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20)"
                .SQL = .SQL & " Where 区分名称 = '内示年月'"
                .SQL = .SQL & "   And 区分タイプ = ''"
                .SQL = .SQL & "   And 値 = '1'"
                .DBExec
            End With
        End If
        
'        If Not NaijiPut Then
'          Call ProcLogPut("内示データ更新異常。＜変換処理＞")
'
'          typText.ERRCODE(HIdx) = "N1"
'          Call TERPut("N1")
'          GoTo ProcHenkan_E1
'       End If
        
    End If
    ' 識別子更新
    If TKT.TNSタイプ = 1 Then
       For gCnt1 = 0 To UBound(tx_Shiki)
           If Trim(tx_Shiki(gCnt1)) = "" Then Exit For
           gStr = TKT.契約先 & TKT.FCODE & tx_Shiki(gCnt1)
           gSL_Select = "Select * From 名称マスタ "
           gSL_Select = gSL_Select & " Where 区分名称 = 'TEXT買掛金'"
           gSL_Select = gSL_Select & "   And 値 = '" & gStr & "'"
           If Not SYMRead(gSL_Select, 1) Then
              Call ShikiBetsuUpd(tx_Shiki(gCnt1))
           End If
       Next gCnt1
        
       Call SyogoTorcd
    End If
    On Error GoTo 0
    '
    B1_TextProc = "変 換 処 理 終 了"
    Call ProcLogPut("変換処理終了 件数=" & tx_InCnt)
    ProcHenkan = True
    Exit Function
ProcHenkan_E1:
    B1_TextProc = "変 換 処 理 異 常 終 了"
    Call ProcLogPut("変換処理異常終了 件数=" & tx_InCnt)
    Exit Function
ProcHenkan_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String
    If Err = 62 Then
       ER_MSG1 = "【 ファイル内容異常 】" & vbCrLf & "入力ファイル = " & mInFName
    Else
       ER_MSG1 = "【 " & Err.Description & " 】"
    End If

    Call ProcLogPut(ER_MSG1 & "   <変換処理>")

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    mErrStr1 = ER_MSG1
    mErrStr2 = ER_MSG2
    
    typText.ERRCODE(HIdx) = "H8"
    Call TERPut("H8", mErrStr1, mErrStr2)
    
End Function
'+-------------------------------+
'+      TEXT管理テーブル更新       +
'+-------------------------------+
Private Sub TEXTUpdate(TErrcd As String, id As Long)
    ' TEXT管理テーブル更新
    With CisDB
        .SQL = "UPDATE TEXT管理テーブル SET "
        If Trim(TErrcd) = "" Then
           .SQL = .SQL & " 最終処理 = '" & mFirstHead & "',"
        End If
        .SQL = .SQL & " 処理件数 = " & tx_SyoriCnt & ","
        .SQL = .SQL & " 処理結果 = '" & RTrim(TErrcd) & "',"
        .SQL = .SQL & " 最終処理日時 = '" & Now() & "'"
        .SQL = .SQL & "  Where FileCD = '" & TKT.FILECD & "'"
        .ConnectNo = 1
        Call .DBExec
    End With
    
    tx_GridObj.TextMatrix(typText.ProcRow(id), tx_ProcDate) = Format(Now(), "YYYY/MM/DD hh:mm")
    tx_GridObj.TextMatrix(typText.ProcRow(id), tx_Kensu) = Format(tx_SyoriCnt, "#,###")
End Sub
'+-------------------------------+
'+
'+      テキスト内容確認
'+
'+-------------------------------+
Private Function TextCheck(EIdx As Long) As Boolean
    TextCheck = False
    
'    Debug.Print tx_DataStr
    
    If Mid(tx_DataStr, 1, 1) = "H" Then      ' ﾍｯﾀﾞｰ件数加算
       If Trim(mFirstHead) = "" Then       ' 最初のﾍｯﾀﾞｰ内容退避
          mFirstHead = Mid(tx_DataStr, 1, 30)
       End If
       mChackHead = Mid(tx_DataStr, 1, 30)
       If Not HeadConvert Then
          typText.ERRCODE(EIdx) = "H9"
          Call TERPut("H9", mErrStr1, mErrStr2)
          Exit Function
       End If
    ' ﾍｯﾀﾞｰ内容ﾁｪｯｸ
       If Not CheckHeder(mErr) Then
          typText.ERRCODE(EIdx) = mErr
          Exit Function
       End If
       tx_HeadCnt = tx_HeadCnt + 1
     End If
    ' ﾃﾞｰﾀ内容ﾁｪｯｸ
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 INSERT START
'     If (TKT.TEXTレイアウト = 6 And TKT.TNSタイプ = 0) Or _
'        (TKT.TEXTレイアウト = 6 And TKT.TNSタイプ = 1) Or _
'        (TKT.TEXTレイアウト = 7 And TKT.TNSタイプ = 1) Then
'========================================================== 2010/02/15 Start
     If (TKT.TEXTレイアウト = 6 And TKT.TNSタイプ = 0) Or _
        (TKT.TEXTレイアウト = 6 And TKT.TNSタイプ = 1) Or _
        (TKT.TEXTレイアウト = 7) Then
        If TKT.TEXTレイアウト = 6 And TKT.TNSタイプ = 0 Then
           tx_HeadHCd = ""
           tx_HeadYmd = "20" & Mid(tx_DataStr, 62, 6)
           tx_HeadSeq = ""
           tx_HeadFCd = TKT.FILECD
           tx_HSiirCd = Trim(Mid(tx_DataStr, 1, 4))
        Else
'           tx_HeadHCd = ""
'           tx_HeadYmd = "20" & Mid(tx_DataStr, 56, 6)
'           tx_HeadSeq = ""
'           tx_HeadFCd = TKT.FILECD
'           tx_HSiirCd = Trim(Mid(tx_DataStr, 1, 4))
           If TKT.TEXTレイアウト = 6 Then
               tx_HeadHCd = ""
               tx_HeadYmd = "20" & Mid(tx_DataStr, 56, 6)
               tx_HeadSeq = ""
               tx_HeadFCd = TKT.FILECD
               tx_HSiirCd = Trim(Mid(tx_DataStr, 1, 4))
            Else
               tx_HeadHCd = ""
               tx_HeadYmd = Mid(tx_DataStr, 157, 6)
               tx_HeadSeq = ""
               tx_HeadFCd = TKT.FILECD
               tx_HSiirCd = RTrim(TKT.契約先)
            End If
        End If
     Else
'========================================================== 2010/02/15 End
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 INSERT END
        If Not CheckData(tx_DataStr, mErr) Then
           typText.ERRCODE(EIdx) = mErr
           Exit Function
        End If
     End If '--------------------------------------- 2008/05/09 ADD
    
     If TKT.TNSタイプ = 0 Then
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 UPDATE START
        'If Mid(tx_DataStr, typConvert.kStart, typConvert.kLeng) = typConvert.kChar Then      ' データ件数加算
        If Mid(tx_DataStr, typConvert.kStart, typConvert.kLeng) = typConvert.kChar Or _
           (TKT.TEXTレイアウト = 6 Or TKT.TEXTレイアウト = 7) Then       ' データ件数加算       '===== 2010/02/15
'           TKT.TEXTレイアウト = 6 Then      ' データ件数加算
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 UPDATE END
           tx_InCnt = tx_InCnt + 1
           Select Case tx_LayoutType
               Case 1  ' 部品納入内示
                    If Not NaijiConVert(mErr) Then
                       typText.ERRCODE(EIdx) = mErr
                       Exit Function
                    End If
           End Select
        End If
    End If
    
    If TKT.TNSタイプ = 1 Then
       If Mid(tx_DataStr, 1, 1) = "D" Then
           tx_InCnt = tx_InCnt + 1
           If TKT.TEXTレイアウト = "5" Then         ' トヨタ車体(701)
              If Not T701ConVertPut(tx_DataStr, mErr) Then
                 typText.ERRCODE(EIdx) = mErr
                 Exit Function
              End If
           End If
           If TKT.TEXTレイアウト = 1 Then         ' トヨタ紡織(060)
              If Not T060ConVertPut(tx_DataStr, mErr) Then
                 typText.ERRCODE(EIdx) = mErr
                 Exit Function
              End If
           End If
           If TKT.TEXTレイアウト = 2 Then         ' トヨタ紡織(070)
              If Not T070ConVertPut(tx_DataStr, mErr) Then
                 typText.ERRCODE(EIdx) = mErr
                 Exit Function
              End If
           End If
           If TKT.TEXTレイアウト = 3 Then         ' トヨタ紡織(080)
              If Not T080ConVertPut(tx_DataStr, mErr) Then
                 typText.ERRCODE(EIdx) = mErr
                 Exit Function
              End If
           End If
           If TKT.TEXTレイアウト = 4 Then         ' トヨタ紡織(090)
              If Not T090ConVertPut(tx_DataStr, mErr) Then
                 typText.ERRCODE(EIdx) = mErr
                 Exit Function
              End If
           End If
       End If
       If TKT.TEXTレイアウト = 0 Then         ' トヨタ(G53)
          If Not TG53ConVertPut(tx_DataStr, mErr) Then
             typText.ERRCODE(EIdx) = mErr
             Exit Function
          End If
       End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 INSERT START
       If TKT.TEXTレイアウト = 6 Then           ' 車体精工(購入)
          If Not T801ConVertPut(tx_DataStr, mErr) Then
             typText.ERRCODE(EIdx) = mErr
             Exit Function
          End If
       End If
       If TKT.TEXTレイアウト = 7 Then           ' 車体精工(支給)
          If Not T802ConVertPut(tx_DataStr, mErr) Then
              typText.ERRCODE(EIdx) = mErr
              Exit Function
          End If
       End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 INSERT END
    End If
'--------------------------------------------------------------------- 2007/03/05 Insert
    '-------- 手配データ
    If TKT.TNSタイプ = 2 Then
       If TKT.TEXTレイアウト = 1 Then         ' トヨタ紡織(020)
          If Not T020ConVertPut(tx_DataStr, mErr) Then
             typText.ERRCODE(EIdx) = mErr
             Exit Function
          End If
       End If
       If TKT.TEXTレイアウト = 2 Then         ' トヨタ車体(112)
          If Not T112ConVertPut(tx_DataStr, mErr) Then
             typText.ERRCODE(EIdx) = mErr
             Exit Function
          End If
       End If
    End If
'--------------------------------------------------------------------- 2007/03/05 Insert End
    If Mid(tx_DataStr, 1, 1) = "T" Then      ' ﾄﾗｰﾗｰ件数加算
       tx_TailCnt = tx_TailCnt + 1
    End If
        
    tx_SyoriCnt = tx_SyoriCnt + 1
             
    TextCheck = True

End Function
'*------------------------------------------------*'
'*      ヘッダー情報よりファイル情報を求める       *
'*------------------------------------------------*'
Public Function HeadConvert() As Boolean
    HeadConvert = False
    
    tx_HeadHCd = Trim(Mid(tx_DataStr, 7, 5))
    tx_HeadYmd = "20" & Mid(tx_DataStr, 15, 6)
    tx_HeadSeq = Trim(Mid(tx_DataStr, 94, 1) & Mid(tx_DataStr, 21, 2))
    tx_HeadFCd = Mid(tx_DataStr, 12, 3) & Trim(Mid(tx_DataStr, 87, 1))
    tx_HSiirCd = Trim(Mid(tx_DataStr, 2, 5))

    If Trim(tx_HeadHCd) = "" Or _
       Trim(tx_HeadYmd) = "" Or _
       Trim(tx_HeadSeq) = "" Or _
       Trim(tx_HeadFCd) = "" Or _
       Trim(tx_HSiirCd) = "" Then
       GoTo HeadConvert_ER
    End If
    HeadConvert = True
    Exit Function
HeadConvert_ER:
    Call ProcLogPut("テキストファイルヘッダー異常")
    mErrStr1 = "          テキストファイルヘッダー異常          "
    mErrStr2 = "            ファイルコード = '" & TKT.FILECD
End Function
'+--------------------------------------+
'+     ﾃｷｽﾄﾌｧｲﾙ内容確認(ﾍｯﾀﾞｰ情報)       +
'+--------------------------------------+
Private Function CheckHeder(CErr As String) As Boolean
    CheckHeder = False

    CErr = ""
    If RTrim(TKT.FCODE) <> tx_HeadFCd Then   ' ﾌｧｲﾙｺｰﾄﾞが異なる
       Call ProcLogPut("ファイルコード異常 対象ｺｰﾄﾞ=" & TKT.FCODE & ",ﾌｧｲﾙｺｰﾄﾞ=" & tx_HeadFCd)
       mErrStr1 = "     ファイルコードが異なる          "
       mErrStr2 = "   対象ｺｰﾄﾞ  = " & StrConv(TKT.FCODE, vbWide) & vbCrLf & vbCrLf & _
                  "   ﾌｧｲﾙｺｰﾄﾞ  = " & StrConv(tx_HeadFCd, vbWide)
       CErr = "H4"
       Call TERPut(CErr, mErrStr1, mErrStr2)
       Exit Function
    End If
    
    If Trim(mCompany) = "" Then
       tx_JisyaChk = False
    End If
    
    If tx_JisyaChk Then
        If RTrim(mCompany) <> tx_HSiirCd Then '自社向けデータでない
           Call ProcLogPut("自社向けデータでない 自社ｺｰﾄﾞ=" & mCompany & ",ﾌｧｲﾙｺｰﾄﾞ=" & tx_HSiirCd)
           mErrStr1 = "     自社向けデータでない              "
           mErrStr2 = "   自社ｺｰﾄﾞ  = " & StrConv(mCompany, vbWide) & vbCrLf & vbCrLf & _
                      "   ﾌｧｲﾙｺｰﾄﾞ  = " & StrConv(tx_HSiirCd, vbWide)

           CErr = "H5"
           Call TERPut(CErr, mErrStr1, mErrStr2)
           Exit Function
        End If
    End If

    If RTrim(mChackHead) = RTrim(TKT.最終処理) Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     処理済みデータです       "
            .MB_MSG(3) = " "
            .MB_MSG(4) = "    再度処理を行いますか？         "
            .MB_Title = "処理確認"
            .MB_Button = OK_CAN
            If Not .MBOX Then
               Call ProcLogPut("処理済みデータ(処理中止)")
               CErr = "H6"
               mErrStr1 = "処理済みデータ(処理中止)"
               Call TERPut(CErr, mErrStr1)
               Exit Function
            End If
            Call ProcLogPut("処理済みデータ(処理継続)")
            mHeadChkFlg = True
       End With
    End If

    CheckHeder = True

End Function
'+--------------------------------------+
'+      ﾃｷｽﾄﾌｧｲﾙ内容確認(データ内容)     +
'+--------------------------------------+
Private Function CheckData(TData As String, CErr As String) As Boolean
    CheckData = False
    CErr = ""
    If mFirstHead = "" And Mid(TData, 1, 1) <> "H" Then
       Call ProcLogPut("テキストデータファイル異常。(ヘッダー情報なし)")
       CErr = "H7"
       mErrStr1 = "     テキストデータファイルが異常です              "
       mErrStr2 = "           (ヘッダー情報なし)                   "
       Call TERPut(CErr, mErrStr1, mErrStr2)
       Exit Function
    End If
    If Mid(TData, 1, 1) <> "H" And Mid(TData, 1, 1) <> "D" And Mid(TData, 1, 1) <> "T" Then
       Call ProcLogPut("テキストデータファイル異常。(区分異常)")
       CErr = "HA"
       mErrStr1 = "     テキストデータファイルが異常です              "
       mErrStr2 = "           (区分異常)                   "
       Call TERPut(CErr, mErrStr1, mErrStr2)
       Exit Function
    End If

    CheckData = True
End Function
'+-------------------------------------+
'+      ファイル名最大を求める
'+-------------------------------------+
Public Function TextMaxFName(LBaseName As String)
    Static wFName()     As String
    Static strFName     As String
    Static FIdx         As Long
    
    If InStr(1, LBaseName, "*") = 0 Then
       TextMaxFName = LBaseName
       Exit Function
    End If
    
    Erase wFName
    ReDim wFName(1)
'
    strFName = Dir(tx_InputDir & LBaseName)      ' 最初のファイル名を返します。
    If Trim(strFName) = "" Then
       TextMaxFName = ""
       Exit Function
    End If
    
    Do While strFName <> ""                ' ループを開始します。
        FIdx = FIdx + 1
        ReDim Preserve wFName(FIdx)
        wFName(FIdx - 1) = strFName

        strFName = Dir                     ' 次のファイル名を返します。
    Loop
'' TEXTﾌｧｲﾙ名分類(降順:最新ﾌｧｲﾙのみ)
    For FIdx = 1 To UBound(wFName) - 1
        If wFName(0) < wFName(FIdx) Then
           wFName(0) = wFName(FIdx)
        End If
    Next FIdx

    TextMaxFName = wFName(0)
End Function
''+--------------------------------------------------+
''+      内示データ　更新
''+--------------------------------------------------+
'Private Function NaijiPut() As Boolean
'    NaijiPut = False
'    On Error GoTo NaijiPut_er:
'
'    NaijiPut = True
'    Exit Function
'NaijiPut_er:
'
'End Function
'+--------------------------------------------------+
'+      識別子更新
'+--------------------------------------------------+
Private Sub ShikiBetsuUpd(Shiki As String)
    Call ItemsClearSYM
    With SYM
        .区分名称 = "TEXT買掛金"
        .区分桁数 = 15
        .作成者 = gTanto
    
        If TKT.契約先 = "7364" And TKT.FCODE = "060" Then
           Select Case Shiki
                Case "15"
                    .値名称 = "副資材(OCR)"
                Case "26"
                    .値名称 = "補給部品納入ｶｰﾄﾞ"
                    .数字1 = 1
                Case "41"
                    .値名称 = "部品・直送支給品(OCR)"
                    .数字1 = 1
                Case "87"
                    .値名称 = "ＫＤ部品(OCR)"
           End Select
        End If
        If TKT.契約先 = "7364" And TKT.FCODE = "070" Then
           Select Case Shiki
                Case "60"
                    .値名称 = "白納品書"
                Case "60"
                    .値名称 = "白納品書"
                Case "61"
                    .値名称 = "ｸﾞﾘｰﾝ納品書"
                Case "62"
                    .値名称 = "即受入返却書"
                Case "63"
                    .値名称 = "トヨタ納品書(黒)"
                Case "64"
                    .値名称 = "トヨタ納品書(赤)"
                Case "65"
                    .値名称 = "訂正伝票(黒)"
                Case "66"
                    .値名称 = "訂正伝票(赤)"
           End Select
          .数字3 = 1
        End If
        If TKT.契約先 = "7364" And (TKT.FCODE = "080" Or TKT.FCODE = "090") Then
           Select Case Shiki
                Case "41"
                    .値名称 = "直送品"
                    .数字2 = 1
                Case "70"
                    .値名称 = "社内出庫"
                    .数字2 = 1
                Case "71"
                    .値名称 = "返却伝票"
                Case "90"
                    .値名称 = "型支払分"
                    .数字3 = 1
           End Select
        End If
        If TKT.契約先 = "7388" Then
           Select Case Shiki
                Case "11"
                    .値名称 = "部品A部品費"
                    .数字1 = 1
                Case "12"
                    .値名称 = "一般仕入"
                    .数字1 = 1
                Case "13"
                    .値名称 = "補給部品部品費"
                    .数字1 = 1
                Case "14"
                    .値名称 = "部品A型費"
                Case "15"
                    .値名称 = "補給部品型費"
                Case "21"
                    .値名称 = "支給部品"
                    .数字2 = 1
           End Select
        End If
        If TKT.契約先 = "1000" Then
           Select Case Shiki
                Case "A"
                    .値名称 = "号口部品"
                    .数字1 = 1
                Case "B"
                    .値名称 = "補給部品"
                    .数字1 = 1
                Case "D"
                    .値名称 = "その他"
                Case "H"
                    .値名称 = "溯り号口"
                    .数字3 = 1
                Case "I"
                    .値名称 = "溯り補給"
                    .数字3 = 1
                Case "J"
                    .値名称 = "クレーム"
                    .数字3 = 1
                Case "M"
                    .値名称 = "産業車両"
                    .数字3 = 1
                Case "C"
                    .値名称 = "産車部品"
                    .数字3 = 1
                Case "N"
                    .値名称 = "車両付属"
                    .数字3 = 1
                Case "R"
                    .値名称 = "型費"
                    .数字3 = 1
                Case "S"
                    .値名称 = "型費訂正"
                    .数字3 = 1
                Case "O"
                    .値名称 = "号口部品(売掛金)"
                    .数字2 = 1
                Case "P"
                    .値名称 = "その他(売掛金)"
           End Select
        End If
        
        .値 = TKT.契約先 & TKT.FCODE & Shiki
        .英数字1 = TKT.契約先
        .英数字2 = TKT.FCODE
        .英数字3 = Shiki
        Call SYMReadChk
    End With
End Sub
'+-------------------------------------------------+
'+          照合対象契約先　出力
'+-------------------------------------------------+
Private Sub SyogoTorcd()
    gSL_Select = "Select 略称 From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & TKT.契約先 & "'"
    gSL_Select = gSL_Select & "   And 取引先区分 = 0 "
    gSL_Select = gSL_Select & "   And 契約先区分 = 1 "
    TRM.略称 = ""
    Call TRMRead(gSL_Select, 1)
    
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "照合対象契約先"
        .区分桁数 = 7
        .値 = TKT.契約先
        .値名称 = RTrim(TRM.略称)
        .作成者 = gTanto
        .検索名 = "処理対象契約先検索"
        .コード名 = "契約先"
        .名称名 = "契 約 先 名"
        .作成者 = gTanto
    End With
    If SYMInitCreate(True) Then
       Call ItemsClearSYM
       With SYM
            .区分名称 = "照合対象契約先"
            .区分桁数 = 7
            .値 = TKT.契約先
            .値名称 = RTrim(TRM.略称)
            .作成者 = gTanto
            Call SYMReadChk
       End With
    End If
End Sub
