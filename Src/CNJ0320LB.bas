Attribute VB_Name = "CNJ0320LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注内示表　印刷
'*
'*   モジュールID  : CNJ0320.Bas
'*
'*   戻値 : なし
'*
'*   作成日  : 2004/03/24  By CIS
'*   変更日  : 2005/02/10  By CIS - 品番取引先マスタの索引処理を追加（直送先用）
'*---------------------------------------------------------------------------------
    Public CisPrt           As New CISFormPrint
    Public TGet_YM          As String
    Public opOutKB          As Integer

    Public PgBk_Torcd       As String       ' ページブレイクキー
    Public PrtK_Hinbn       As String       ' 明細行出力キー    ( 品番 )
    Public PrtK_Torcd       As String       ' 明細行出力キー    ( 手配先 )
    Public PrtK_Ukeir       As String       ' 明細行出力キー    ( 受入 )

'*---------------------------------------------------* ▽▽▽ 2005/02/10 Insert Start ▽▽▽
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@      □□□ 品番取引先マスタの索引処理 □□□
'@
'@          >> 索引項目：納入先受入
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function DBSelect_HTM(Hinbn As String, _
                             Torcd As String, _
                             Ukeir As String, _
                           Hinmoku As Byte, _
                            NTorcd As String, Optional DBNo As Integer = 1) As String
'
    DBSelect_HTM = ""
'>> 直送先の判定
    If RTrim$(NTorcd) <> "" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "select htm.納入先受入"
        gSL_Select = gSL_Select & "  from 品番取引先マスタ htm"
        gSL_Select = gSL_Select & " where htm.品番 = '" & RTrim$(Hinbn) & "'"
        gSL_Select = gSL_Select & "   and htm.取引先 = '" & RTrim$(Torcd) & "'"
        gSL_Select = gSL_Select & "   and htm.受入 = '" & RTrim$(Ukeir) & "'"
        gSL_Select = gSL_Select & "   and htm.品目 = 1"
        gSL_Select = gSL_Select & "   and htm.納入先 = '" & RTrim$(NTorcd) & "'"
    '>> 索引
        If HTMRead(gSL_Select, 1, DBNo) Then DBSelect_HTM = RTrim$(HTM.納入先受入)
    End If
'
End Function
'*---------------------------------------------------* △△△ 2005/02/10 Insert End   △△△

