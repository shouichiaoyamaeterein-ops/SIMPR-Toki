Attribute VB_Name = "CKNZ0220Bas"
Option Explicit
'+----------------------------------------------------------+
'+
'+       モジュール名  :かんばん発行用　共通モジュール
'+       モジュールID  :CKNZ0220Bas.BAS
'+       作成日        :2003/12/25   CIS
'+       更新日        :2006/02/02   By CIS - INIファイル-クライアント対応（CommPort）
'+       更新日        :
'+
'+----------------------------------------------------------+
    Global kn_Call          As String
    Global kn_PrtKB         As Byte
    Global kn_PrtSY         As Byte
    Global Kn_PCnt          As Long

    Global Kn_Jisya         As String
    Global Kn_JisyaNM       As String
    Global Kn_JisyaKJNM     As String
    
    Global Kn_Rtn           As Boolean
    Global Kn_TestFlg       As Boolean
    Global kn_PortNo        As Integer
    Global kn_Setting       As String
    Global kn_Hand          As Integer
    Global kn_232cLog       As Boolean
    
    Global Kn_Page          As Boolean
    Global Kn_PageS         As Long
    Global Kn_PageE         As Long
    Global Kn_Hopa          As Integer
    
    Global kBas_Proc        As String       ' ダミー
    
    Global SaveControlName  As Control
    Global XOnOffFlg        As Boolean
    
    Type T_ライン
        始点X         As Long
        始点Y         As Long
        終点X         As Long
        終点Y         As Long
        種別          As Long
        幅            As Long
        丸み          As Long
    End Type
    
    Global Kn_Line()  As T_ライン
    
    Type T_ビットマップ
        始点X         As Long
        始点Y         As Long
        横倍          As Byte
        縦倍          As Byte
        フォント      As String
        文字間        As Integer
        回転          As Byte
        背景          As String
        フィールド    As String
        固有設定      As String
        固定          As String
        桁数          As Long
        少数          As Byte
        編集          As String
        位置          As Byte
    End Type
    
    Global Kn_Bit()     As T_ビットマップ
    
    Type T_アウトライン
        始点X         As Long
        始点Y         As Long
        文字幅        As Long
        文字高        As Long
        フォント      As String
        文字間        As Integer
        回転          As Byte
        背景          As String
        フィールド    As String
        固有設定      As String
        固定          As String
        桁数          As Long
        少数          As Byte
        編集          As String
        位置          As Byte
    End Type
    
    Global Kn_Out()     As T_アウトライン
    
    Type T_バーコード
        始点X         As Long
        始点Y         As Long
        フィールド    As String
        桁数          As Long
    End Type
    Global Kn_Bar()     As T_バーコード
    
    Dim knId                 As Long
    Dim wFixStr         As String * 30
    
'******************************************'
'      現在の設定値を初期ﾌｧｲﾙより求める
'******************************************'
Public Sub KnGetComm()

    With CisFun
        .INI_Section = "Apparatus"
        .INI_FileID = gIniExe & gIniName
        .INI_Key = "3"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "TX-4370"    '
           .PutIni
        End If
        
        .INI_Section = .INI_String
'-----------------------------------------------< 2006/02/02 MNT START >
'        .INI_FileID = gIniExe & gIniName
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2006/02/02 MNT  END  >
        ' --- デバッグ用(かんばん出力)
        If gDemoPrt <> 0 Then
           Kn_TestFlg = True
        Else
           Kn_TestFlg = False
        End If
        ' --- デバッグ用(232C出力)
        .INI_Default = ""
        .INI_Key = "RS232CLog"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "No"
           .PutIni
        End If
        If UCase(.INI_String) = "Y" Or UCase(.INI_String) = "YES" Then
           kn_232cLog = True
        Else
           kn_232cLog = False
        End If
        ' '----- CommPort獲得 -----
        .INI_Default = ""
        .INI_Key = "CommPort"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "1"
           .PutIni
        End If
        kn_PortNo = CisFun.Val2(.INI_String)
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2006/02/02 ADD  END  >
        ' '----- 通信速度獲得 -----
        .INI_Default = ""
        .INI_Key = "Setting"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "9600,N,8,1"
           .PutIni
        End If
        kn_Setting = .INI_String
        ' '----- HandShaking獲得 -----
        .INI_Default = ""
        .INI_Key = "HandShaking"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "2"
           .PutIni
        End If
        kn_Hand = CisFun.Val2(.INI_String)
        
        ' かんばん出力ファイル設定
        .Log_PutFlg = Kn_TestFlg Or kn_232cLog
        .Log_FileInitial = "KanWrite"
        .Log_FilePath = gIniExe & "Log\"
        .Log_DataHD = ""
        .Log_DataHDTime = あり
        .Log_DelDay = 3

        .LogPut ""
    End With
    
    
End Sub
'*--------------------------------------------------------------------------------
'*
'*   ｺﾐｭﾆｹｰｼｮﾝ(MSComm) ｺﾝﾄﾛｰﾙ初期処理
'*
'*   モジュールID  : InitComm
'*
'*          ( INIﾌｧｲﾙ使用の為、『IniGet』後実行 )
'*--------------------------------------------------------------------------------
Public Function InitComm(IComm As Control) As Boolean
    InitComm = False
    
    With IComm
        .CommPort = kn_PortNo
        .Settings = kn_Setting
        .Handshaking = kn_Hand
        .InputLen = 0
        .RThreshold = 1
        .SThreshold = 1
        .OutBufferSize = 4096
    End With

    Set SaveControlName = IComm
End Function
'*--------------------------------------------------------------------------------
'*
'*   フォーマット情報取得
'*
'*   モジュールID  : FormatGet
'*
'*--------------------------------------------------------------------------------
Public Function FormatGet() As Boolean
    FormatGet = False
    ReDim Kn_Line(1)
    '---(ﾗｲﾝ ﾌｫｰﾏｯﾄ)
    gSL_Select = "Select * From かんばん設定詳細テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & kn_PrtSY
    gSL_Select = gSL_Select & "   And 設定区分 = 1"
    gSL_Select = gSL_Select & "   And 表示 = ''"
    gSL_Select = gSL_Select & " Order By 設定連番"
    If KSSRead(gSL_Select) Then
        ReDim Kn_Line(CisDB.RecordCount)
        knId = 0
        Do Until Not KSS_RDSTS
            With Kn_Line(knId)
                .始点X = KSS.始点X
                .始点Y = KSS.始点Y
                .終点X = KSS.終点X
                .終点Y = KSS.終点Y
                .種別 = KSS.種別
                .幅 = KSS.幅
                .丸み = KSS.丸み
            End With
            knId = knId + 1
            
            Call KSSReadNext
        Loop
    End If
    
    Call KSSClose
    
    '---(ﾋﾞｯﾄﾏｯﾌﾟﾌｫﾝﾄ ﾌｫｰﾏｯﾄ)
    ReDim Kn_Bit(1)
    gSL_Select = "Select * From かんばん設定詳細テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & kn_PrtSY
    gSL_Select = gSL_Select & "   And 設定区分 = 2"
    gSL_Select = gSL_Select & "   And 表示 = ''"
    gSL_Select = gSL_Select & " Order By 設定連番"
    If KSSRead(gSL_Select) Then
       ReDim Kn_Bit(CisDB.RecordCount)
        knId = 0
        Do Until Not KSS_RDSTS
            With Kn_Bit(knId)
                .始点X = KSS.始点X
                .始点Y = KSS.始点Y
                .横倍 = KSS.横倍
                .縦倍 = KSS.縦倍
                .フォント = KSS.フォント
                .文字間 = KSS.文字間
                .回転 = KSS.回転
                .背景 = KSS.背景
                .固有設定 = KSS.固有設定
                .フィールド = KSS.フィールド
                .固定 = KSS.固定
                .桁数 = KSS.桁数
                .少数 = KSS.少数
                .編集 = KSS.編集
                .位置 = KSS.位置
            End With
            
            knId = knId + 1
            
            Call KSSReadNext
        Loop
    End If
    Call KSSClose
    '---(ｱｳﾄﾗｲﾝﾌｫﾝﾄ ﾌｫｰﾏｯﾄ)
    ReDim Kn_Out(1)
    gSL_Select = "Select * From かんばん設定詳細テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & kn_PrtSY
    gSL_Select = gSL_Select & "   And 設定区分 = 3"
    gSL_Select = gSL_Select & "   And 表示 = ''"
    gSL_Select = gSL_Select & " Order By 設定連番"
    If KSSRead(gSL_Select) Then
       ReDim Kn_Out(CisDB.RecordCount)
        knId = 0
        Do Until Not KSS_RDSTS
            With Kn_Out(knId)
                .始点X = KSS.始点X
                .始点Y = KSS.始点Y
                .文字高 = KSS.文字高
                .文字幅 = KSS.文字幅
                .フォント = KSS.フォント
                .文字間 = KSS.文字間
                .回転 = KSS.回転
                .背景 = KSS.背景
                .固有設定 = KSS.固有設定
                .フィールド = KSS.フィールド
                .固定 = KSS.固定
                .桁数 = KSS.桁数
                .少数 = KSS.少数
                .編集 = KSS.編集
                .位置 = KSS.位置
            End With
            
            knId = knId + 1
            
            Call KSSReadNext
        Loop
    End If
    Call KSSClose
    
    '---(ﾊﾞｰｺｰﾄﾞ ﾌｫｰﾏｯﾄ)
    ReDim Kn_Bar(1)
    gSL_Select = "Select * From かんばん設定詳細テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & kn_PrtSY
    gSL_Select = gSL_Select & "   And 設定区分 = 4"
    gSL_Select = gSL_Select & "   And 表示 = ''"
    gSL_Select = gSL_Select & " Order By 設定連番"
    If KSSRead(gSL_Select) Then
       ReDim Kn_Bar(CisDB.RecordCount)
       knId = 0
        Do Until Not KSS_RDSTS
            With Kn_Bar(knId)
                .始点X = KSS.始点X
                .始点Y = KSS.始点Y
                .フィールド = KSS.フィールド
                .桁数 = KSS.桁数
            End With
            
            knId = knId + 1
            
            Call KSSReadNext
        Loop
    End If
    Call KSSClose
    
    FormatGet = True
End Function
'*--------------------------------------------------------------------------------
'*
'*   かんばん書き出し
'*
'*   モジュールID  : FormatRSSet
'*
'*--------------------------------------------------------------------------------
Public Function FormatRSSet() As Boolean
    Dim wPutData    As String
    
    FormatRSSet = False
    ' 発行
    If KST.ホッパ = 1 Then
       wPutData = "F" & Format(KHT.StHNo, "0")
       wPutData = wPutData & Format(KHT.EdHNo, "0")
       wPutData = wPutData & Format(KHT.Stack1, "0")
       wPutData = wPutData & Format(KHT.Stack2, "0")
       Call RSPut(wPutData)
    End If
    
    '---(ﾗﾍﾞﾙｻｲｽﾞ設定)
    wPutData = "D" & Format(KST.ラベルピッチ, "0000")
    wPutData = wPutData & "," & Format(KST.有効印字幅, "0000")
    wPutData = wPutData & "," & Format(KST.有効印字長, "0000")
    Call RSPut(wPutData)
    '---(初期化)
    wPutData = "C"
    Call RSPut(wPutData)
    '---(ﾗｲﾝ ﾌｫｰﾏｯﾄ)
    For knId = 0 To UBound(Kn_Line) - 1
        If (Kn_Line(knId).始点X) <> 0 Then
            wPutData = "LC;" & Format(Kn_Line(knId).始点X, "0000")
            wPutData = wPutData & "," & Format(Kn_Line(knId).始点Y, "0000")
            wPutData = wPutData & "," & Format(Kn_Line(knId).終点X, "0000")
            wPutData = wPutData & "," & Format(Kn_Line(knId).終点Y, "0000")
            wPutData = wPutData & "," & Format(Kn_Line(knId).種別, "0")
            wPutData = wPutData & "," & Format(Kn_Line(knId).幅, "0")
            If Kn_Line(knId).丸み <> 0 Then
               wPutData = wPutData & "," & Format(Kn_Line(knId).丸み, "000")
            End If
            Call RSPut(wPutData)
        End If
    Next knId
    '---(ﾋﾞｯﾄﾏｯﾌﾟﾌｫﾝﾄ ﾌｫｰﾏｯﾄ)
    For knId = 0 To UBound(Kn_Bit) - 1
        With Kn_Bit(knId)
            If .始点X <> 0 Then
                wPutData = "PC" & Format(knId + 1, "000") & ";"
                wPutData = wPutData & "" & Format(.始点X, "0000")
                wPutData = wPutData & "," & Format(.始点Y, "0000")
                wPutData = wPutData & "," & Format(.横倍, "00")
                wPutData = wPutData & "," & Format(.縦倍, "00")
                wPutData = wPutData & "," & .フォント
                Select Case .文字間
                    Case Is > 0: wPutData = wPutData & ",+" & Format(.文字間, "00")
                    Case Is < 0: wPutData = wPutData & ",-" & Format(.文字間 * -1, "00")
                End Select
                If .回転 <> 0 Then
                   wPutData = wPutData & "," & Format(.回転, "00")
                Else
                   wPutData = wPutData & ",00"
                End If
                wPutData = wPutData & "," & .背景
                If .位置 <> 0 And .位置 <> 3 Then
                   wPutData = wPutData & ",P" & Format(.位置, "0")
                End If
                
                Call RSPut(wPutData)
            End If
        End With
        
        
    Next knId
    '---(ｱｳﾄﾗｲﾝﾌｫﾝﾄ ﾌｫｰﾏｯﾄ)
    For knId = 0 To UBound(Kn_Out) - 1
        With Kn_Out(knId)
            If .始点X <> 0 Then
                wPutData = "PV" & Format(knId + 1, "00") & ";"
                wPutData = wPutData & "" & Format(.始点X, "0000")
                wPutData = wPutData & "," & Format(.始点Y, "0000")
                wPutData = wPutData & "," & Format(.文字幅, "0000")
                wPutData = wPutData & "," & Format(.文字高, "0000")
                wPutData = wPutData & "," & .フォント
                Select Case .文字間
                    Case Is > 0: wPutData = wPutData & ",+" & Format(.文字間, "000")
                    Case Is < 0: wPutData = wPutData & ",-" & Format(.文字間 * -1, "000")
                End Select
                If .回転 <> 0 Then
                   wPutData = wPutData & "," & Format(.回転, "00")
                Else
                   wPutData = wPutData & ",00"
                End If
                wPutData = wPutData & "," & .背景
                If .位置 <> 0 And .位置 <> 3 Then
                   wPutData = wPutData & ",P" & Format(.位置, "0")
                End If
                
                Call RSPut(wPutData)
            End If
        End With
        
        
    Next knId
    
    '---(ﾊﾞｰｺｰﾄﾞ ﾌｫｰﾏｯﾄ)
    For knId = 0 To UBound(Kn_Bar) - 1
        With Kn_Bar(knId)
            If .始点X <> 0 Then
                wPutData = "XB" & Format(knId + 1, "00") & ";"
                wPutData = wPutData & "" & Format(.始点X, "0000")
                wPutData = wPutData & "," & Format(.始点Y, "0000")
                wPutData = wPutData & "," & KST.バー種類
                wPutData = wPutData & "," & KST.誤り訂正
                wPutData = wPutData & "," & Format(KST.セル幅, "00")
                wPutData = wPutData & "," & KST.モード
                wPutData = wPutData & "," & Format(KST.回転, "0")
                If KST.モデル <> 0 Then
                   wPutData = wPutData & "," & Format(KST.モデル, "0")
                End If
                Call RSPut(wPutData)
            End If
        End With
        
    Next knId
    '--------------------------------------'
    '       印  字  デ  ー  タ             '
    '--------------------------------------'
    '---(ﾋﾞｯﾄﾏｯﾌﾟﾌｫﾝﾄ ﾃﾞｰﾀ)
    For knId = 0 To UBound(Kn_Bit) - 1
        If Kn_Bit(knId).桁数 <> 0 Then
            With Kn_Bit(knId)
                wPutData = "RC" & Format(knId + 1, "000") & ";"
                gStr = ""
                If Trim(.固有設定) = "" Then
                    If .フィールド <> "" Then
                       gStr = DataGet(.フィールド)
                       Select Case .編集
                            Case "1": gStr = Format(CisFun.Val2(gStr), String(.桁数, "0"))
                            Case "2": gStr = Format(CisFun.Val2(gStr), "#,##0")
                            Case "3": gStr = Format(CisFun.Val2(gStr), "#,###")
                            Case "4": gStr = Format(CisFun.Val2(gStr), "#,##0." & String(.少数, "0"))
                            Case "5"
                                If CisFun.Val2(gStr) <> 0 Then
                                   gStr = Format(CisFun.Val2(gStr), "#,##0." & String(.少数, "0"))
                                Else
                                   gStr = ""
                                End If
                       End Select
'                       If .位置 = 3 Then
'                          gStr = Right(gStr, .桁数)
'                       End If
                    Else
                       gStr = CisFun.Mid2(.固定, 1, .桁数)
                    End If
                Else
                    Select Case .固有設定
                        Case "1": gStr = Format(Date, "YYYY/MM/DD")
                        Case "2": gStr = Format(Date, "YY/MM/DD")
                        Case "3": gStr = Format(Date, "YYYYMMDD")
                                  gStr = Mid(gStr, 1, 4) & "年" & Mid(gStr, 5, 2) & "月" & Mid(gStr, 6, 2) & "日"
                    End Select
                End If
                
                If .桁数 > CisFun.Len2(gStr) Then
                   If .位置 <> 3 Then
                      gStr = gStr & Space(.桁数 - CisFun.Len2(gStr))
                   Else
                      gStr = Space(.桁数 - CisFun.Len2(gStr)) & gStr
                   End If
                End If
                wPutData = wPutData & CisFun.Mid2(gStr, 1, .桁数)
            End With
            
            Call RSPut(wPutData)
        End If
    Next knId
    '---(ｱｳﾄﾗｲﾝﾌｫﾝﾄ ﾃﾞｰﾀ)
    For knId = 0 To UBound(Kn_Out) - 1
        If Kn_Out(knId).桁数 <> 0 Then
            With Kn_Out(knId)
                wPutData = "RV" & Format(knId + 1, "00") & ";"
                gStr = ""
                If Trim(.固有設定) = "" Then
                    If .フィールド <> "" Then
                       gStr = DataGet(.フィールド)
                       Select Case .編集
                            Case "1": gStr = Format(CisFun.Val2(gStr), String(.桁数, "0"))
                            Case "2": gStr = Format(CisFun.Val2(gStr), "#,##0")
                            Case "3": gStr = Format(CisFun.Val2(gStr), "#,###")
                            Case "4": gStr = Format(CisFun.Val2(gStr), "#,##0." & String(.少数, "0"))
                            Case "5"
                                If CisFun.Val2(gStr) <> 0 Then
                                   gStr = Format(CisFun.Val2(gStr), "#,##0." & String(.少数, "0"))
                                Else
                                   gStr = ""
                                End If
                       End Select
'                       If .位置 = 3 Then
'                          gStr = Right(gStr, .桁数)
'                       End If
                    Else
                       gStr = CisFun.Mid2(.固定, 1, .桁数)
                    End If
                Else
                    Select Case .固有設定
                        Case "1": gStr = Format(Date, "YYYY/MM/DD")
                        Case "2": gStr = Format(Date, "YY/MM/DD")
                        Case "3": gStr = Format(Date, "YYYYMMDD")
                                  gStr = Mid(gStr, 1, 4) & "年" & Mid(gStr, 5, 2) & "月" & Mid(gStr, 6, 2) & "日"
                    End Select
                End If
                
                If .桁数 > CisFun.Len2(gStr) Then
                   If .位置 <> 3 Then
                      gStr = gStr & Space(.桁数 - CisFun.Len2(gStr))
                   Else
                      gStr = Space(.桁数 - CisFun.Len2(gStr)) & gStr
                   End If
                End If
                wPutData = wPutData & CisFun.Mid2(gStr, 1, .桁数)
            End With
            
            Call RSPut(wPutData)
        End If
    Next knId
    '---(ﾊﾞｰｺｰﾄﾞ ﾃﾞｰﾀ)
    For knId = 0 To UBound(Kn_Bar) - 1
        If Kn_Bar(knId).桁数 <> 0 Then
            With Kn_Bar(knId)
                wPutData = "RB" & Format(knId + 1, "00") & ";"
                gStr = ""
                If .フィールド <> "" Then
                   gStr = DataGet(.フィールド)
                End If
                If .桁数 > CisFun.Len2(gStr) Then
                   gStr = gStr & Space(.桁数 - CisFun.Len2(gStr))
                End If
                wPutData = wPutData & CisFun.Mid2(gStr, 1, .桁数)
            End With
            
            Call RSPut(wPutData)
        End If
    Next knId
    FormatRSSet = True
End Function
'*--------------------------------------------------------------------------------
'*
'*   かんばん発行
'*
'*   モジュールID  : KanbanPrint
'*
'*--------------------------------------------------------------------------------
Public Sub KanbanPrint()
    Dim PrtStr      As String
'    ' 発行
'    If KST.ホッパ = 1 Then
'       PrtStr = "F" & Format(KHT.StHNo, "0")
'       PrtStr = PrtStr & Format(KHT.EdHNo, "0")
'       PrtStr = PrtStr & Format(KHT.Stack1, "0")
'       PrtStr = PrtStr & Format(KHT.Stack2, "0")
'       Call RSPut(PrtStr)
'    End If
'
    PrtStr = "XS;I,0001," & Format(KST.カット間隔, "000")
    PrtStr = PrtStr & Format(KST.センサ, "0") & "C"
    PrtStr = PrtStr & Format(KST.印字スピード, "0")
    PrtStr = PrtStr & Format(KST.リボン, "0")
    PrtStr = PrtStr & "10"
    
    Call RSPut(PrtStr)

End Sub
'*--------------------------------------------------------------------------------
'*
'*   RS-232c データ出力
'*
'*   モジュールID  : RSPut(IComm As MSComm, IData As String)
'*
'*      最終更新日 ： 2003/12/25
'*--------------------------------------------------------------------------------
Public Sub RSPut(IData As String)
    
    IData = Chr$(&H1B) & IData & Chr$(&HA) & Chr$(&H0)
    
    If Not Kn_TestFlg Then
        Do
            DoEvents
            If Not XOnOffFlg Then

                SaveControlName.Output = IData

                If kn_232cLog Then
                   CisFun.LogPut (IData)
                End If
                DoEvents
                Exit Do
            End If
        Loop
    Else
'       * テスト用テキストファイル書き出し *
        Call CisFun.LogPut(IData)
    End If
End Sub
Private Function DataGet(Filed As String) As String
    With KHT
        Select Case Filed
            Case "かんばん種類":          DataGet = .かんばん種類
            Case "品番":                  DataGet = .品番
            Case "取引先":                DataGet = .取引先
            Case "受入":                  DataGet = .受入
            Case "発行NO":                DataGet = .発行NO
            Case "背番号":                DataGet = .背番号
            Case "発行区分":              DataGet = .発行区分
            Case "表示品番":              DataGet = .表示品番
            Case "品番14":                DataGet = .品番14
            Case "品番30":                DataGet = .品番30
            Case "品名":                  DataGet = .品名
            Case "受入名":                DataGet = .受入名
            Case "収容器":                DataGet = .収容器
            Case "収容数":                DataGet = .収容数
            Case "手配区分":              DataGet = .手配区分
            Case "手配区分名":            DataGet = .手配区分名
            Case "アドレス":              DataGet = .アドレス
            Case "取引先CD":              DataGet = .取引先CD
            Case "取引先名":              DataGet = .取引先名
            Case "取引先名1":             DataGet = .取引先名1
            Case "取引先名2":             DataGet = .取引先名2
            Case "取引先工場名":          DataGet = .取引先工場名
            Case "取引先工場名1":         DataGet = .取引先工場名1
            Case "取引先工場名2":         DataGet = .取引先工場名2
            Case "納入先固定":            DataGet = .納入先固定
            Case "納入先CD":              DataGet = .納入先CD
            Case "納入先名":              DataGet = .納入先名
            Case "納入先名1":             DataGet = .納入先名1
            Case "納入先名2":             DataGet = .納入先名2
            Case "納入先工場名":          DataGet = .納入先工場名
            Case "納入先受入":            DataGet = .納入先受入
            Case "メーカ":                DataGet = .メーカ
            Case "メーカ名":              DataGet = .メーカ名
            Case "メーカ名1":             DataGet = .メーカ名1
            Case "メーカ名2":             DataGet = .メーカ名2
            Case "メーカ受入":            DataGet = .メーカ受入
            Case "サイクル":              DataGet = .サイクル
            Case "置場":                  DataGet = .置場
            Case "車型":                  DataGet = .車型
            Case "自社名1":               DataGet = .自社名1
            Case "自社名2":               DataGet = .自社名2
            Case "工程":                  DataGet = .工程
            Case "工程名":                DataGet = .工程名
            Case "次工程":                DataGet = .次工程
            Case "次工程名":              DataGet = .次工程名
            Case "基準工数":              DataGet = .基準工数
            Case "生産ロット":            DataGet = .生産ロット
            Case "基準生産日":            DataGet = .基準生産日
            Case "材質":                  DataGet = .材質
            Case "材質18":                DataGet = .材質18
            Case "寸法区分":              DataGet = .寸法区分
            Case "板厚":                  DataGet = .板厚
            Case "幅":                    DataGet = .幅
            Case "長さ":                  DataGet = .長さ
            Case "材料名":                DataGet = .材料名
            Case "使用品番":              DataGet = .使用品番
            Case "単位名":                DataGet = .単位名
            Case "手配単位名":            DataGet = .手配単位名
            Case "実数区分":              DataGet = .実数区分
            Case "材種コード":            DataGet = .材種コード
            Case "備考":                  DataGet = .備考
            Case "区分":                  DataGet = .区分
            Case "納入時間1":             DataGet = .納入時間1
            Case "納入時間2":             DataGet = .納入時間2
            Case "納入時間3":             DataGet = .納入時間3
            Case "納入時間4":             DataGet = .納入時間4
            Case "納入時間5":             DataGet = .納入時間5
            Case "納入時間6":             DataGet = .納入時間6
            Case "納入時間7":             DataGet = .納入時間7
            Case "納入時間8":             DataGet = .納入時間8
            Case "得意先背番1":           DataGet = .得意先背番1
            Case "得意先背番2":           DataGet = .得意先背番2
            Case "得意先背番3":           DataGet = .得意先背番3
            Case "得意先背番4":           DataGet = .得意先背番4
            Case "得意先工場名1":         DataGet = .得意先工場名1
            Case "得意先工場名2":         DataGet = .得意先工場名2
            Case "得意先工場名3":         DataGet = .得意先工場名3
            Case "得意先工場名4":         DataGet = .得意先工場名4
            Case "QR":                    DataGet = QR_OutCreate
            Case "発行済":                DataGet = .発行済
            Case "作成者":                DataGet = .作成者
        End Select
    End With
End Function
