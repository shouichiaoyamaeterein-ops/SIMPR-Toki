VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Begin VB.Form CXT0020 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "SIMPR-A"
   ClientHeight    =   7290
   ClientLeft      =   5220
   ClientTop       =   2895
   ClientWidth     =   10110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   7290
   ScaleWidth      =   10110
   Begin Cis3D_v60.CIS3D B1_TextProc 
      Height          =   885
      Left            =   210
      Top             =   660
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1561
      BackColor       =   0
      ForeColor       =   65535
      Caption         =   "デ ー タ 作 成 処 理 開 始"
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
      Height          =   405
      Left            =   210
      ScaleHeight     =   345
      ScaleWidth      =   9615
      TabIndex        =   2
      Top             =   1590
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
      ItemData        =   "CXT0020.frx":0000
      Left            =   210
      List            =   "CXT0020.frx":0002
      TabIndex        =   0
      Top             =   2070
      Width           =   9645
   End
End
Attribute VB_Name = "CXT0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  ＴＮＳ送信データ作成処理
'**       フォームID    :  CXT0020
'**       処理概要      :
'**
'**       作  成  日    :  2005/09/02  By CIS
'**       変  更  日    :  2006/02/08  By CIS  自動化対応
'**       変  更  日    :  2006/03/07  By CIS  ファミック用カスタマイズ
'**       変  更  日    :  2007/02/08  By CIS  車体支給品入荷送信追加(C04)
'**       変  更  日    :  2007/02/09  By CIS  TNSｼｽﾃﾑ起動ﾁｪｯｸ追加
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim CisFrm          As CISFormContorl
 
    Dim cIDX            As Long
    Dim cIDX2           As Long
    
    Dim mCrtDataFlg     As Boolean
    
    Dim mTaskReturn     As Byte         ' TNSｼｽﾃﾑ起動確認       2007/02/09 Insert
    Dim mSendErr        As Boolean                             '2007/02/09 Insert
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
    B1_TextFMEI = ""
    ProcList.Clear
    Set Ts_TextProcObj = B1_TextProc
    Set Ts_ProcListObj = ProcList
    Set Ts_BarObj = B1_Gage
    
    Set Ts_FromObj = New CISFormContorl
    Set Ts_FromObj.ThisForm = Me
'   #-----------------------------#
'   #     TNS送信　初期処理       #
'   #-----------------------------#
    If Ts_ProcType = 0 Then
       Call TnsSend_Init
    End If
'
    Dummy.Left = -1000
'   #-------------------#
'   # 画面中央位置表示  #
'   #-------------------#
'---------------------------------------------------------- 2006/02/08 Insert
    If Ts_Auto And Not Ts_AutoDisp Then
       'Me.WindowState = 0
       Me.Left = Me.Width * -1
       Me.Top = Me.Height * -1
    Else
       If Ts_Auto And Ts_AutoDisp Then
          Me.Left = 0
          Me.Top = 0
       Else
          Call CisFrm.Frm_Center(False, x1024y768)
       End If
    End If
'---------------------------------------------------------- 2006/02/08 Insert End

'---------------------------------------------------------- 2006/02/08 Delete
    'Call CISFrm.Frm_Center(False, x1024y768)
'---------------------------------------------------------- 2006/02/08 Delete End

End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》           +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = True
    
    DoEvents
    
    Dummy.SetFocus
    
    Ts_ProcErr = True
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
    Call TNSSend_LogPut("E")
    
    Me.Hide
End Sub
'**********************************
'*     テキスト変換・累積処理       *
'**********************************
Private Sub ReturnPress()
    Dim wDelDate        As String
    Me.MousePointer = 11
    
    On Error GoTo ReturnPress_Ed
    
'+------------------------------------------+
'+
'+      処     理     開      始
'+
'+------------------------------------------+
    For cIDX = 0 To UBound(Ts_SndFCode) - 1

        Ts_Torcd = Ts_SndTorcd(cIDX)
        Ts_FCode = Ts_SndFCode(cIDX)
        Ts_KTorcd = Ts_SndKTorcd(cIDX)
        Ts_SendKb = Ts_SndSendKb(cIDX)
        
        Ts_SendProc = True
    
        If Not Start_TSRInsert(Ts_Torcd, Ts_FCode, Ts_ReSendNo) Then
           Call TNSSend_MsgDisp("", "送信履歴　初期化エラー", 0, "ER")
           Exit Sub
        End If
    '+-----------> TNS送信設定テーブル 読込み
        gSL_Select = "Select TS.*,TR.略称 送信先名 from TNS送信設定テーブル TS "
    '+-------------------------------------------------------------------------- 2007/02/09 Delete
'        gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR "
'        gSL_Select = gSL_Select & "     ON   TR.取引先 = TS.送信先"
'        gSL_Select = gSL_Select & "     And  TR.取引先区分 = 0"
    '+-------------------------------------------------------------------------- 2007/02/09 Delete End
        
'    '+------------------------------------------------------------------- 2007/02/09 Insert
'        gSL_Select = gSL_Select & " Left Outer Join (Select MIN(取引先) 取引先,略称"
'        gSL_Select = gSL_Select & "                              From 取引先マスタ TR"
'        gSL_Select = gSL_Select & "                    Where 取引先区分 = 0"
'        gSL_Select = gSL_Select & "                       And  契約先区分 = 1"
'        gSL_Select = gSL_Select & "                 Group By 取引先,略称 ) TR"
'        gSL_Select = gSL_Select & "              ON TR.取引先 = TS.送信先"
'    '+------------------------------------------------------------------- 2007/02/09 Insert End
    '+------------------------------------------------------------------- 2007/12/27 Insert
        gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR"
        gSL_Select = gSL_Select & "   ON TR.取引先CD = TS.送信先"
        gSL_Select = gSL_Select & "   And  契約先区分 = 1"
    '+------------------------------------------------------------------- 2007/02/27 Insert End
        
        gSL_Select = gSL_Select & " Where TS.送信先 = '" & Ts_KTorcd & "'"
        gSL_Select = gSL_Select & "   And TS.FCODE = '" & Ts_FCode & "'"
        If Ts_SendKb = 0 Then
            gSL_Select = gSL_Select & "   And TS.代行送信先 = ''"                           '2007/12/27 ADD
        Else
            gSL_Select = gSL_Select & "   And TS.代行送信先 = '" & RTrim(Ts_Torcd) & "'"    '2007/12/27 ADD
        End If
        gSL_Select = gSL_Select & "   And TS.送信先区分 = " & Ts_SendKb & ""            '2007/12/27 ADD
        If Not TSSRead(gSL_Select, 1) Then
           Ts_SendProc = False
           gStr = "*Err TNS送信設定テーブル　未登録 送信先=" & Ts_Torcd & " 契約先=" & Ts_KTorcd & vbCrLf
           gStr = gStr & "                           FCODE = " & Ts_FCode
           Call TNSSend_LogPut("", gStr)
           Call Middle_TSRUpDate("01", "TNS送信設定　未登録")
        Else
           If Ts_ReSendNo = 0 Then
              TSR.実行区分 = 0         ' ﾃﾞｰﾀ作成
           Else
              TSR.実行区分 = 4         ' 再処理(作成)
           End If
        '+------------------------------------------------------------------- 2007/02/09 Insert
           B1_TextFMEI = TSS.送信先名 & "   【" & TSS.ファイル名称 & "】"
           If Ts_EDIstartFlg Then
              If TSS.TNSシステム起動 = 1 Then
                 mTaskReturn = CisFun.TackCheck("Edisys01.exe")
                 If mTaskReturn <> 0 Then
                    Ts_SendProc = False
                    gStr = "*Err TNSｼｽﾃﾑ未起動"
                    Call TNSSend_MsgDisp("", gStr, 10, "ER")
                    Call Middle_TSRUpDate("31", "TNSｼｽﾃﾑ未起動")
                                     
                    If Not Ts_Auto Then
                       With CisFun
                         .MB_Lines = 5
                         .MB_MSG(2) = "     TNSｼｽﾃﾑ未起動です       "
                         .MB_MSG(3) = " "
                         .MB_MSG(4) = "     処理を行いますか？         "
                         .MB_Title = "処理確認"
                         .MB_Button = OK_CAN
                         If Not .MBOX Then
                            gStr = "TNSｼｽﾃﾑ未起動(処理中止)"
                            Call TNSSend_MsgDisp("", gStr, 10, "ER")
                            Ts_SendProc = False
                         Else
                            gStr = "TNSｼｽﾃﾑ未起動(処理継続)"
                            Call TNSSend_MsgDisp("", gStr, 10, "ER")
                            Ts_SendProc = True
                         End If
                       End With
                    End If
                 Else
                    gStr = "★★★★★★ ＴＮＳシステム起動済"
                    Call TNSSend_MsgDisp("", gStr, 10)
                 End If
              End If
           End If
        '+------------------------------------------------------------------- 2007/02/09 Insert End
           If Ts_SendProc Then                                        '----- 2007/02/09 Insert
           
           Call Middle_TSRUpDate("11", "履歴情報削除エラー")
           
           B1_TextFMEI = TSS.送信先名 & "   【" & TSS.ファイル名称 & "】"
            
           gStr = "【" & TSS.ファイル名称 & "】 処理開始 "
           Call TNSSend_MsgDisp(gStr, gStr)
           gStr = "送信先 = " & Ts_Torcd & TSS.送信先名 & "  ﾌｧｲﾙｺｰﾄﾞ = " & Ts_FCode
           Call TNSSend_MsgDisp("", gStr, 20)
    
           gStr = "   履 歴 情 報 削 除 "
           Call TNSSend_MsgDisp("", gStr)
            
           Ts_SendFName = TSS.TNS送信ファイル名
            
            ' バックアップテキスト & 履歴 削除
            If Not TNSSend_BackKill(Ts_Torcd, Ts_KTorcd, Ts_FCode, TSS.削除日数) Then
               Ts_SendProc = False
            End If
    
        '+-----------> TNS送信用　データ　作成
            Call Middle_TSRUpDate("12", "TNS送信用ﾃﾞｰﾀ作成ｴﾗｰ")
            mCrtDataFlg = False
            C04_HinErrCnt = 0
            ' 直送支給品データ
            If TSS.TNSタイプ = 1 Or TSS.TNSタイプ = 2 Then
               If Not TNSSEND_C04 Then
                  Ts_SendProc = False
               Else
                  mCrtDataFlg = True
               End If
            End If
''    '------------------------------------------------------------------------------- 2007/02/08 Insert
''            ' 直送支給品データ(車体(C04)
''            If TSS.TNSタイプ = 2 Then
''               mR07Err_Flg = 0
''               If Not TNSSEND_C04 Then
''                  Ts_SendProc = False
''               Else
''                  mCrtDataFlg = True
''               End If
''            End If
''    '------------------------------------------------------------------------------- 2007/02/08 Insert End
           End If                                                      '----- 2007/02/09 Insert
    
        '+-----------> TNS送信(TNS通信ﾊﾟｯｹｰｼﾞｲﾝｽﾄｰﾙ済のみ)
            If Ts_ProcType = 3 Then
               TSS.送信 = 1
            End If
            If Ts_ProcType <> 4 And _
               mCrtDataFlg And Ts_TNSPakegeFlg And TSS.送信 = 1 And TSS.ゼロ件データ = 0 Then
               If Ts_ReSendNo = 0 Then
                  TSR.実行区分 = 1         ' 送信
               Else
                  TSR.実行区分 = 3         ' 再処理(作成・送信)
               End If
               TSR.送信開始日時 = Now()
               Call Middle_TSRUpDate("13", "TNS送信ｴﾗｰ")
            
               gStr = "【　　TNS送信処理　開始　　】"
               Call TNSSend_MsgDisp("", gStr, 10)
               
               Ts_TNSFName = Ts_SendPath & Ts_SendFName
               Ts_Shikibetsu = RTrim(TSS.識別子)
               Ts_Atesaki = RTrim(TSS.TNS宛先)
               
               gStr = "送信元ﾌｧｲﾙ名=" & Ts_TNSFName
               Call TNSSend_MsgDisp("", gStr, 15)
               gStr = "識別子=" & Ts_Shikibetsu & "/TNS宛先=" & Ts_Atesaki
               Call TNSSend_MsgDisp("", gStr, 15)
               
               If Not TNS_File_Send_Api Then
                  Ts_SendProc = False
                  gStr = "【　　TNS送信処理　異常終了　　】ｴﾗｰｺｰﾄﾞ = " & Ts_TNSReturn
                  Call TNSSend_MsgDisp("", gStr, 10)
                  TSR.ERR = Ts_TNSReturn
               Else
                  gStr = "【　　TNS送信処理　終了　　】"
                  Call TNSSend_MsgDisp("", gStr, 10)
               End If
            
               TSR.送信終了日時 = Now()
               Call Middle_TSRUpDate("UPD")  ' 履歴更新
            End If
            
        '+-----------> 終了処理
            If Ts_SendProc Then
'               Call Middle_TSRUpDate("00")   ' 正常終了
               If C04_HinErrCnt = 0 Then
                   Call Middle_TSRUpDate("00")   ' 正常終了
                   Ts_ProcErr = False
               Else
                   Call Middle_TSRUpDate("95", "【020】品番未登録有")  ' 正常終了だけど、品番未登録有(FAMIC) 2006/03/07
               End If
            Else
                Ts_ProcErr = True
            End If
        '+-----------> グリッド更新(送信)
            If Not Ts_Grid Is Nothing Then
               With Ts_Grid
                    .TextMatrix(Ts_SndRow(cIDX), Ts_ColProcData) = Format(TSR.開始日時, "YYYY/MM/DD hh:mm")
                    .TextMatrix(Ts_SndRow(cIDX), Ts_ColSendCnt) = Format(TSR.送信件数, "#,###")
                    .TextMatrix(Ts_SndRow(cIDX), Ts_ColDataCnt) = Format(TSR.データ件数, "#,###")
                    .TextMatrix(Ts_SndRow(cIDX), Ts_ColRNo) = TSR.履歴管理NO
                    .TextMatrix(Ts_SndRow(cIDX), Ts_ColResultCD) = TSR.処理結果
                    .TextMatrix(Ts_SndRow(cIDX), Ts_ColErrCode) = TSR.ERR
                    
                    .TextMatrix(Ts_SndRow(cIDX), Ts_ColSelect) = ""
                    
                    If Trim(TSR.処理結果) > "10" Then
                       .TextMatrix(Ts_SndRow(cIDX), Ts_ColResult) = "ｴﾗｰ"
                       Ts_ProcErr = True
                    Else
                       .TextMatrix(Ts_SndRow(cIDX), Ts_ColResult) = ""
                    End If
                    .Row = Ts_SndRow(cIDX)
                    For gInt = 0 To .Cols - 1
                        .Col = gInt
                        If gInt = 0 Then
                          .CellBackColor = .BackColorFixed
                        Else
                          .CellBackColor = .BackColor
                        End If
                    Next gInt
               End With
            End If
        End If
        
    Next cIDX

'    Ts_ProcErr = False
ReturnPress_Ed:
    Me.MousePointer = 1
    
    Call EndProc
    
End Sub
