Attribute VB_Name = "KanbanQR_Bas"
Option Explicit
'+----------------------------------------------------------+
'+
'+       モジュール名  :かんばんQR用　共通モジュール
'+       モジュールID  :KanbanQR_Bas.BAS
'+       作成日        :2004/01/19   CIS
'+       更新日        :2008/04/10   CIS 背番号８桁対応
'+
'+----------------------------------------------------------+
    Dim wFixStr         As String * 30

    Global Qr_Jisya     As String
    Global Qr_HMoto     As String
    
    Public Type QR_Area
        バーコード種類   As Integer
        ステータス       As Integer
        発行元          As String
        自社コード      As String
        品番            As String
        品目            As Byte
        取引先          As String
        受入            As String
        手配区分        As String
        版数            As Byte
        発行日          As String
        背番号          As String
        連番            As Integer
        収容数          As Long
        材料収容数      As Currency
        サイクル        As String
        納入先          As String
        納入先受入      As String
        工程            As String
        次工程          As String
        材料管理番号    As Long
        編集品番        As String
        手配区分名      As String
        アドレス        As String
        板厚            As Currency
        幅              As Currency
        長さ            As Currency
        寸法区分        As Byte
    End Type
    
    Public QRArea       As QR_Area
    
    Public QRChk_SimplrA    As String
    Public QRChk_SimplrABar As String
'*----------------------------------------------*'
'*      ワークエリアクリア  ( QRデータ作成 )
'*----------------------------------------------*'
Public Sub QRAreaClear()
    With QRArea
        .ステータス = 0
        .発行元 = ""
        .自社コード = ""
        .品番 = ""
        .品目 = 0
        .取引先 = ""
        .受入 = ""
        .手配区分 = ""
        .版数 = 0
        .発行日 = ""
        .背番号 = ""
        .連番 = 0
        .収容数 = 0
        .材料収容数 = 0
        .サイクル = ""
        .納入先 = ""
        .納入先受入 = ""
        .工程 = ""
        .次工程 = ""
        .材料管理番号 = 0
        .編集品番 = ""
        .手配区分名 = ""
        .アドレス = ""
        .板厚 = 0
        .幅 = 0
        .長さ = 0
        .寸法区分 = 0
    End With
End Sub
'+--------------------------------------+
'+      QRデータ作成
'+--------------------------------------+
Public Function QR_OutCreate() As String
    Dim wQRData     As String
    QR_OutCreate = ""
    wQRData = ""
    With QRArea
         .発行元 = Qr_HMoto
         .自社コード = Qr_Jisya
         wFixStr = .発行元
         wQRData = wQRData & Mid(wFixStr, 1, 7)
         wFixStr = .自社コード
         wQRData = wQRData & Mid(wFixStr, 1, 5)
         wFixStr = .品番
         wQRData = wQRData & Mid(wFixStr, 1, 25)
         wFixStr = .品目
         wQRData = wQRData & Format(.品目, "0")
         wFixStr = .取引先
         wQRData = wQRData & Mid(wFixStr, 1, 7)
         wFixStr = .受入
         wQRData = wQRData & Mid(wFixStr, 1, 2)
         wFixStr = .手配区分
         wQRData = wQRData & Mid(wFixStr, 1, 1)
         wFixStr = Format(.版数, "00")
         wQRData = wQRData & Mid(wFixStr, 1, 2)
         If Trim(.発行日) = "" Then .発行日 = Format(Date, "YYYYMMDD")
         wFixStr = .発行日
         wQRData = wQRData & Mid(wFixStr, 1, 8)
         wFixStr = .背番号
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE START
         'wQRData = wQRData & Mid(wFixStr, 1, 5)
         wQRData = wQRData & Mid(wFixStr, 1, 8)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE END
         wFixStr = Format(.連番, "000")
         wQRData = wQRData & Mid(wFixStr, 1, 3)
         wFixStr = Format(.収容数, "000000")
         wQRData = wQRData & Mid(wFixStr, 1, 6)
         .サイクル = Replace(.サイクル, "-", "")
         .サイクル = Replace(.サイクル, " ", "0")
         wFixStr = .サイクル
         wQRData = wQRData & Mid(wFixStr, 1, 5)
         wFixStr = .納入先
         wQRData = wQRData & Mid(wFixStr, 1, 7)
         wFixStr = .納入先受入
         wQRData = wQRData & Mid(wFixStr, 1, 2)
         wFixStr = .工程
         wQRData = wQRData & Mid(wFixStr, 1, 3)
         wFixStr = .次工程
         wQRData = wQRData & Mid(wFixStr, 1, 3)
         wFixStr = Format(.材料管理番号, "00000000")
         wQRData = wQRData & Mid(wFixStr, 1, 8)
         wFixStr = .編集品番
         wQRData = wQRData & Mid(wFixStr, 1, 25)
         wFixStr = .手配区分名
         wQRData = wQRData & Mid(wFixStr, 1, 2)
         wFixStr = .アドレス
         wQRData = wQRData & Mid(wFixStr, 1, 7)
         wFixStr = Format(.板厚, "0000.00")
         wQRData = wQRData & Mid(wFixStr, 1, 7)
         wFixStr = Format(.幅, "0000.00")
         wQRData = wQRData & Mid(wFixStr, 1, 7)
         wFixStr = Format(.長さ, "0000.00")
         wQRData = wQRData & Mid(wFixStr, 1, 7)
         wFixStr = Format(.寸法区分, "0")
         wQRData = wQRData & Mid(wFixStr, 1, 1)
         wFixStr = Format(.材料収容数, "000000.000")
         wQRData = wQRData & Mid(wFixStr, 1, 10)
'         ' 空き
'         wFixStr = ""
'         wQRData = wQRData & Mid(wFixStr, 1, 2)
    End With

    QR_OutCreate = wQRData
End Function
'+----------------------------------+
'+      QRかんばん情報変換(Simpr-A)
'+----------------------------------+
Public Function QR_SimprA(QRDATA As String) As Boolean
    QR_SimprA = False
    
    If Not RTrim(CisFun.Mid2(QRDATA, 1, 7)) = Qr_HMoto Then
        Exit Function
    End If
    
    Call QRAreaClear
    With QRArea
        .バーコード種類 = 0
        .発行元 = RTrim(CisFun.Mid2(QRDATA, 1, 7))
        .自社コード = RTrim(CisFun.Mid2(QRDATA, 8, 5))
        .品番 = RTrim(CisFun.Mid2(QRDATA, 13, 25))
        .品目 = CisFun.Val2(CisFun.Val2(CisFun.Mid2(QRDATA, 38, 1)))
        .取引先 = RTrim(CisFun.Mid2(QRDATA, 39, 7))
        .受入 = RTrim(CisFun.Mid2(QRDATA, 46, 2))
        .手配区分 = RTrim(CisFun.Mid2(QRDATA, 48, 1))
        .版数 = CisFun.Val2(CisFun.Mid2(QRDATA, 49, 2))
        .発行日 = RTrim(CisFun.Mid2(QRDATA, 51, 8))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE START
'        .背番号 = RTrim(CisFun.Mid2(QRDATA, 59, 5))
'        .連番 = CisFun.Val2(CisFun.Mid2(QRDATA, 64, 3))
'        .収容数 = CisFun.Val2(CisFun.Mid2(QRDATA, 67, 6))
'        .サイクル = RTrim(CisFun.Mid2(QRDATA, 73, 5))
'        .納入先 = RTrim(CisFun.Mid2(QRDATA, 78, 7))
'        .納入先受入 = RTrim(CisFun.Mid2(QRDATA, 85, 2))
'        .工程 = RTrim(CisFun.Mid2(QRDATA, 87, 3))
'        .次工程 = RTrim(CisFun.Mid2(QRDATA, 90, 3))
'        .材料管理番号 = CisFun.Val2(CisFun.Mid2(QRDATA, 93, 8))
'        .編集品番 = RTrim(CisFun.Mid2(QRDATA, 101, 25))
'        .手配区分名 = RTrim(CisFun.Mid2(QRDATA, 126, 4))
'        .アドレス = RTrim(CisFun.Mid2(QRDATA, 130, 7))
'        .板厚 = CisFun.Val2(CisFun.Mid2(QRDATA, 137, 7))
'        .幅 = CisFun.Val2(CisFun.Mid2(QRDATA, 144, 7))
'        .長さ = CisFun.Val2(CisFun.Mid2(QRDATA, 151, 7))
'        .寸法区分 = CisFun.Val2(CisFun.Mid2(QRDATA, 158, 1))
'        .材料収容数 = CisFun.Val2(CisFun.Mid2(QRDATA, 159, 10))
        .背番号 = RTrim(CisFun.Mid2(QRDATA, 59, 8))
        .連番 = CisFun.Val2(CisFun.Mid2(QRDATA, 67, 3))
        .収容数 = CisFun.Val2(CisFun.Mid2(QRDATA, 70, 6))
        .サイクル = RTrim(CisFun.Mid2(QRDATA, 76, 5))
        .納入先 = RTrim(CisFun.Mid2(QRDATA, 81, 7))
        .納入先受入 = RTrim(CisFun.Mid2(QRDATA, 88, 2))
        .工程 = RTrim(CisFun.Mid2(QRDATA, 90, 3))
        .次工程 = RTrim(CisFun.Mid2(QRDATA, 93, 3))
        .材料管理番号 = CisFun.Val2(CisFun.Mid2(QRDATA, 96, 8))
        .編集品番 = RTrim(CisFun.Mid2(QRDATA, 104, 25))
        .手配区分名 = RTrim(CisFun.Mid2(QRDATA, 129, 4))
        .アドレス = RTrim(CisFun.Mid2(QRDATA, 133, 7))
        .板厚 = CisFun.Val2(CisFun.Mid2(QRDATA, 140, 7))
        .幅 = CisFun.Val2(CisFun.Mid2(QRDATA, 147, 7))
        .長さ = CisFun.Val2(CisFun.Mid2(QRDATA, 154, 7))
        .寸法区分 = CisFun.Val2(CisFun.Mid2(QRDATA, 161, 1))
        .材料収容数 = CisFun.Val2(CisFun.Mid2(QRDATA, 162, 10))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE END
    End With
    
    QR_SimprA = True
    
End Function
'+----------------------------------+
'+      QRかんばん情報変換(Simpr-A)
'+----------------------------------+
Public Function QR_SimprABar(QRDATA As String) As Boolean
    Dim wSt         As Integer
    Dim wEd         As Integer
    Dim wCnt        As Integer
    Dim wReadFlg    As Integer
    
    QR_SimprABar = False
    
    If Not IsNumeric(QRDATA) Then Exit Function
    
    wReadFlg = 0
    
    wSt = 0: wEd = 1
    If kBas_Proc = "" Then
       wSt = 0: wEd = 0
    End If
    If kBas_Proc = "Z" Then
       wSt = 1: wEd = 1
    End If
    If kBas_Proc = "S" Then
       wSt = 0: wEd = 1
    End If
    
    For wCnt = wSt To wEd
        With CisDB
            Select Case wCnt
                Case 0
                    .SQL = "SELECT * From 品番取引先マスタ "
                    .SQL = .SQL & " Where Cast(アドレス As Int) = " & CisFun.Val2(Mid(QRDATA, 2, 5))
                    wReadFlg = 1
                Case 1
                    .SQL = "SELECT * From 材料取引先マスタ "
                    .SQL = .SQL & " Where Cast(アドレス As Int) = " & CisFun.Val2(Mid(QRDATA, 2, 5))
                    wReadFlg = 2
                Case Else
                    Exit Function
            End Select
            
            If .DBRead Then
               Exit For
            End If
            .Recordset.Close
            wReadFlg = 0
        End With
    Next wCnt
        
    If wReadFlg = 0 Then Exit Function
        
    Call QRAreaClear
    With QRArea
        .バーコード種類 = 1
        .発行元 = Qr_HMoto
        .自社コード = Qr_Jisya
        .連番 = CisFun.Val2(Mid(QRDATA, 8, 3))
        
        If Not IsNull(CisDB.Recordset("取引先")) Then .取引先 = RTrim(CisDB.Recordset("取引先"))
        If Not IsNull(CisDB.Recordset("受入")) Then .受入 = RTrim(CisDB.Recordset("受入"))
        If Not IsNull(CisDB.Recordset("手配区分")) Then .手配区分 = RTrim(CisDB.Recordset("手配区分"))
        If Not IsNull(CisDB.Recordset("背番号")) Then .背番号 = RTrim(CisDB.Recordset("背番号"))
        If Not IsNull(CisDB.Recordset("サイクル")) Then .サイクル = RTrim(CisDB.Recordset("サイクル"))
        If Not IsNull(CisDB.Recordset("アドレス")) Then .アドレス = RTrim(CisDB.Recordset("アドレス"))
        
        If wReadFlg = 1 Then
            If Not IsNull(CisDB.Recordset("品番")) Then .品番 = RTrim(CisDB.Recordset("品番"))
            If Not IsNull(CisDB.Recordset("品目")) Then .品目 = CisFun.Val2(CisDB.Recordset("品目"))
            If Not IsNull(CisDB.Recordset("収容数1")) Then .収容数 = CisFun.Val2(CisDB.Recordset("収容数1"))
        End If
        If wReadFlg = 2 Then
            If Not IsNull(CisDB.Recordset("材質")) Then .品番 = RTrim(CisDB.Recordset("材質"))
            If Not IsNull(CisDB.Recordset("材料管理番号")) Then .材料管理番号 = CisFun.Val2(CisDB.Recordset("材料管理番号"))
            If Not IsNull(CisDB.Recordset("収容数1")) Then .材料収容数 = CisFun.Val2(CisDB.Recordset("収容数1"))
            .品目 = 2
        End If
    End With
    
    On Error Resume Next
    CisDB.Recordset.Close
    On Error GoTo 0
    
    QR_SimprABar = True
    
End Function

