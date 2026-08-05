Attribute VB_Name = "SysMenuINI"
Option Explicit
'**********************************************************'
'**
'**       作  成  日    :
'**       変  更  日    :   2006/07/06  By CIS 業務ﾒﾆｭｰ切替
'**
'**********************************************************'

    Global gMM_Inf(15)       As String       'ﾒﾆｭｰ タイトル
    Global gMM_InfKey(15)    As String       'ﾒﾆｭｰ タイトル(ｷｰ)
    Global gMS_Inf(15, 3)    As String       'ｻﾌﾞﾒﾆｭｰ タイトル
    Global gMTitle(15, 60)   As String       ' 実行部  ﾀｲﾄﾙ
    Global gMExeKB(15, 60)   As String       ' 実行部  実行区分
    Global gMExe(15, 60)     As String       ' 実行部  Exe名
    Global gMExeNo(15, 20)   As Integer      ' 実行部  Exe実行番号（ﾌｧﾝｸｼｮﾝ対応）
    Global gMS_Key(3)        As String
    
    Global gMId         As Integer
    Global gMMax        As Integer
    
    Global gFHM_No      As Byte
    Global gFMM_No      As Byte
    Global gFT_Text     As String
    Global gEdiExeDir   As String

'--  --  --  --  --  --  --  --  --  < 2006/07/06 ADD >
'
    Global gK_Last          As Byte         ' 業務ﾒﾆｭｰ最終表示№
    Global gKind_Cnt        As Byte         ' 業務ﾒﾆｭｰ数(№)
    Global gK_Title()       As String       ' 業務ﾒﾆｭｰタイトル
    Global gK_Text()        As String       ' 業務ﾒﾆｭｰテキスト( Full Path )

    Global gK_MNSelKB       As Byte         ' ﾒﾆｭｰ選択区分( 0:選択、1:ﾒﾆｭｰ登録 )
    Global gK_MNSelRtn      As Byte         ' ﾒﾆｭｰ選択ﾘﾀｰﾝ( 1:選択した )
    Global gK_EditKB        As Byte         ' 編集処理区分( 1:内容変更行った )

'
'    Global gKBColr_Menu     As Long         ' 業務ﾒﾆｭｰカラー
'    Global gKBColr_Sel      As Long         ' ボタン選択カラー
'    Global gKBColr_Rel      As Long         ' ボタン未選択カラー
'
'
'    Global Const gc_Menu_Colr       As Long = &HFFC0C0      ' ﾒﾆｭｰｶﾗｰ       ( 薄青 )
'    Global Const gc_SBtn_Colr       As Long = &HC0FFC0      ' ﾎﾞﾀﾝ選択ｶﾗｰ   ( 薄緑 )
'    Global Const gc_RBtn_Colr       As Long = &HFF8080      ' ﾎﾞﾀﾝ未選択ｶﾗｰ ( 中青 )
'
'--  --  --  --  --  --  --  --  --  < 2006/07/06 ADD END >

'   * Windows の終了 *
    Public Const EWX_LOGOFF = 0 '現在のﾕｰｻﾞｰのｱﾌﾟﾘｹｰｼｮﾝをｸﾛｰｽﾞし、ﾛｸﾞｵﾌする
    Public Const EWX_SHUTDOWN = 1 'ｼｽﾃﾑをｼｬｯﾄﾀﾞｳﾝ後、電源ｵﾌ可能状態にする
    Public Const EWX_REBOOT = 2 'ｼｽﾃﾑをｼｬｯﾄﾀﾞｳﾝ後、再起動
    Public Const EWX_FORCE = 4 '応答のないﾌﾟﾛｾｽを強制終了する

    Declare Function ExitWindowsEx Lib "USER32" (ByVal uFlags&, ByVal dwReserved&) As Long

' 関連 起動用
    Declare Function ShellExecute Lib _
        "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

    Public Const SW_SHOWNORMAL = 1
    
    Dim mIdx        As Long
    Dim mIdx2       As Long
    Dim mStr        As String
    Dim mSubMax     As Long
'/*-----------------------------------*/
'/*     メニュー情報 テーブルセット     */
'/*-----------------------------------*/
Public Function FromText()
    FromText = False

    Static wFT_Str          As String
    
'   < メニューテキスト取得 >
    With CisFun
        .INI_Section = "MenuText"
        .INI_Key = .WSID
        .INI_Default = ""
        .INI_FileID = gIniExe & gIniName
        
        .GetIni
        gFT_Text = .INI_String
    
        If RTrim$(gFT_Text) = "" Then
            gFT_Text = gIniExe & "Menu.Txt"
            
            .INI_String = gFT_Text
            .PutIni
        End If
    End With
''--  --  --  --  --  --  --  --  --  < 2006/07/06 ADD >
''
''   * メニューテキスト獲得 *
    If gK_Last <> 0 Then
        With CisFun
            .INI_FileID = gIniExe & gIniName
            .INI_Section = .WSID
            .INI_Default = ""
            .INI_Key = "MenuText" & Format(gK_Last, "-00;-00;#")
            .GetIni

            gFT_Text = RTrim$(.INI_String)
            If RTrim$(gFT_Text) = "" Then
                gFT_Text = gIniExe & "Menu.Txt"
            End If
        End With
    End If
''
''--  --  --  --  --  --  --  --  --  < 2006/07/06 ADD END >
    
    On Error GoTo FromText_Err

'    ' テキスト分割用ワーク
    Static wFT_RText        As String
    Static wFT_Split()      As String
    Static wFT_BText(10)    As String
    Static wFT_I            As Integer
    
    Static wFT_MNo          As Byte
    Static wFT_ENo          As Byte
    Static wFT_SNo          As Byte
    Static ReadFlg          As Boolean
    
    Erase gMM_Inf
    Erase gMM_InfKey
    Erase gMS_Inf
    Erase gMTitle
    Erase gMExeKB
    Erase gMExe
    Erase gMS_Key
    
    ' テキスト読込
    gMMax = 0
    
    With CisFun
        .INI_Section = "Menu"
        .INI_Key = "HCNT"
        .INI_Default = ""
        .INI_FileID = gFT_Text
        
        .GetIni
        gMMax = CisFun.Val2(.INI_String)
        
        .INI_Key = "SUBKEY"
        .INI_Default = ""
        
        .GetIni
        gMS_Key(0) = Mid(.INI_String, 1, 1)
        gMS_Key(1) = Mid(.INI_String, 2, 1)
        gMS_Key(2) = Mid(.INI_String, 3, 1)
        
        ' ﾒﾆｭｰﾀｲﾄﾙ取得
        For wFT_I = 0 To gMMax
            .INI_Section = "Menu"
            .INI_Key = Format(wFT_I, "0")
            .INI_Default = ""
            .INI_FileID = gFT_Text
            
            .GetIni
            mStr = Trim(.INI_String)
            If mStr <> "" Then
               wFT_Split = Split(mStr, ",")
               If UBound(wFT_Split) > 0 Then
                  gMM_Inf(wFT_I) = wFT_Split(0)
                  gMM_InfKey(wFT_I) = wFT_Split(1)
                  
'                  .INI_Section = wFT_Split(1)
'                  .INI_Key = "H"
'                  .INI_Default = ""
'                  .INI_FileID = gFT_Text
'                  .INI_String = wFT_Split(0)
'
'                  .PutIni
               End If
            End If
            
        Next wFT_I
    End With
    
    ' テキストファイル オープン
    Static wFT_FileNo       As Integer

    wFT_FileNo = FreeFile

    wFT_Str = "テキストファイルオープン"
    Open gFT_Text For Input Access Read As wFT_FileNo
'
    ReadFlg = False
    
    wFT_Str = "テーブルセット"
    Do While Not EOF(wFT_FileNo)
'
        Line Input #wFT_FileNo, wFT_RText
'
        Erase wFT_Split:    Erase wFT_BText
'
        If Mid(wFT_RText, 1, 1) <> ";" And Trim(wFT_RText) <> "" Then

           If Mid(wFT_RText, 1, 3) = "99=" Then
              ReadFlg = True
           End If
            
           If ReadFlg Then
              If Mid(wFT_RText, 1, 1) = "[" Then
                 mStr = Replace(Trim(wFT_RText), "[", "")
                 mStr = Replace(mStr, "]", "")
                 wFT_MNo = 99
                 For mIdx = 0 To UBound(gMM_Inf) - 1
                     If gMM_InfKey(mIdx) = mStr Then
                        wFT_MNo = mIdx
                        Exit For
                     End If
                 Next mIdx
              End If
              If wFT_MNo <> "99" Then
                If Mid(wFT_RText, 1, 1) = "M" Then
                   wFT_SNo = CisFun.Val2(Mid(wFT_RText, 2, 1))
                   If wFT_MNo < UBound(gMM_Inf) And wFT_SNo < UBound(gMS_Inf, 2) Then
                       gMS_Inf(wFT_MNo, wFT_SNo) = Mid(wFT_RText, 4)
                   End If
                End If
                If Mid(wFT_RText, 1, 1) = "B" Then
                   mStr = Mid(wFT_RText, 6)
                   wFT_Split = Split(mStr, ",")
                   If UBound(wFT_Split) > 1 Then
                      wFT_ENo = CisFun.Val2(Mid(wFT_RText, 3, 2))
                      wFT_SNo = CisFun.Val2(Mid(wFT_RText, 2, 1))
                      If wFT_SNo <> 0 And wFT_SNo < 3 Then
                         wFT_ENo = wFT_ENo + wFT_SNo * 20
                      End If
                      gMTitle(wFT_MNo, wFT_ENo) = wFT_Split(0)
                      gMExeKB(wFT_MNo, wFT_ENo) = wFT_Split(1)
                      gMExe(wFT_MNo, wFT_ENo) = wFT_Split(2)
                   End If
                End If
             End If
           End If
        End If
    Loop

    Close #wFT_FileNo
    
    If Trim(gMS_Key(0)) = "" Then gMS_Key(0) = "Z"
    If Trim(gMS_Key(1)) = "" Then gMS_Key(1) = "X"
    If Trim(gMS_Key(2)) = "" Then gMS_Key(2) = "C"
    
    wFT_Str = ""
    FromText = True
    Exit Function
FromText_Err:
    If Err = 53 Then
        gStr = "初期導入作業が正しく行われていません。" & vbCr & vbCr & _
                 Space(10) & "契約処理を行ってから、再度実行して下さい。"
        MsgBox gStr, vbCritical, ""
        Exit Function
    End If
    
    Call CisFun.ErrorBox(wFT_Str)
End Function
