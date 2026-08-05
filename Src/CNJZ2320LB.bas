Attribute VB_Name = "CNJZ2320LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注内示表　印刷
'*
'*   モジュールID  : CNJ2320.Bas
'*
'*   戻値 : なし
'*
'*   作成日  : 2005/06/06  By CIS
'*   変更日  : 2005/02/10  By CIS - 品番取引先マスタの索引処理を追加（直送先用）
'*---------------------------------------------------------------------------------
    Global CisPrt           As New CISFormPrint
    Public TGet_NJYM        As String           ' 内示年月
    Public TGet_YM          As String
    Public TGet_YMD         As String
    Public opOutKB          As Integer

    Public PgBk_Torcd       As String       ' ページブレイクキー
    Public PrtK_Torcd       As String       ' 明細行出力キー    ( 手配先 )
    Public PrtK_Zainm       As String       ' 明細行出力キー    ( 材料名 )
'*-------------------------------------------------------*
'*  稼働日数取得(社内＆発注分カレンダーマスタ)
'*-------------------------------------------------------*
Public Sub GetKadou()
    ItemsClearCLM
    With CLM
        .取引先区分 = "1"                   '発注
        .取引先 = ""                        '社内カレンダー
        .年 = Mid(RTrim(TGet_YM), 1, 4)
        .月 = Mid(RTrim(TGet_YM), 5, 2)
        .受入 = ""                          '社内カレンダー
        If Not CLMRead("", 1, 1) Then Exit Sub
    End With
End Sub
