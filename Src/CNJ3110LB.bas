Attribute VB_Name = "CNJ3110LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   内示表　印刷（キノックス向けレイアウト）
'*
'*   モジュールID  : CNJ3110.Bas
'*
'*   戻値 : なし
'*
'*   作成日  :  2005/06/06   By CIS
'*   変更日  :  2006/01/30              帳票ヘッド部に『受入』追加(改ページ条件)
'*   変更日  :
'*
'*---------------------------------------------------------------------------------
    Public CisPrt           As New CISFormPrint
    Public CNJ3110_Naiji    As String

    Public TGet_YM          As String
    Public TGet_YMD         As String
    Public opOutKB          As Integer

    Public PgBk_Torcd       As String       ' ページブレイクキー

'< 2006/01/30   ADD START >--------------------------------------------------------------------------------
    Public PgBk_Ukeir       As String       ' ページブレイクキー
'< 2006/01/30   ADD  END  >--------------------------------------------------------------------------------
    
    Public PgBk_Syake       As String       ' ページブレイクキー
    Public PrtK_Hinbn       As String       ' 明細行出力キー    ( 品番 )
    Public PrtK_Torcd       As String       ' 明細行出力キー    ( 手配先 )
    Public PrtK_Ukeir       As String       ' 明細行出力キー    ( 受入 )
    Public PrtK_Seban       As String       ' 明細行出力キー    ( 背番号 )
    Public PrtK_Syuyo       As Long         ' 明細行出力キー    ( 収容数 )
    Public PrtK_Syake       As String       ' 明細行出力キー    ( 車型 )
    Public wNengetu(2)      As String       ' 対象年月
    Public wKadou(2)        As Integer      ' 月別稼働日数
'*-------------------------------------------------------*
'*  稼働日数取得(社内＆受注分カレンダーマスタ)
'*-------------------------------------------------------*
Public Sub GetKadou(ParYYYYMM As String)
    ItemsClearCLM
    
    '契約先・受注分の読込み
    With CLM
        .取引先区分 = "0"                   '受注
        .取引先 = RTrim$(NTT.契約先)
        .年 = Mid(RTrim(ParYYYYMM), 1, 4)
        .月 = Mid(RTrim(ParYYYYMM), 5, 2)
        .受入 = RTrim$(NTT.受入)
        If Not CLMRead("", 1, 1) Then
            '社内・受注分の読込み
            ItemsClearCLM
            .取引先区分 = "0"                   '受注
            .取引先 = ""                        '社内カレンダー
            .年 = Mid(RTrim(ParYYYYMM), 1, 4)
            .月 = Mid(RTrim(ParYYYYMM), 5, 2)
            .受入 = ""                          '社内カレンダー
            If Not CLMRead("", 1, 1) Then Exit Sub
        End If
    End With
End Sub
