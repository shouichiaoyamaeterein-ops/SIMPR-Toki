Attribute VB_Name = "æˆøæ•ÏŠ·ƒ}ƒXƒ^b’è"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) d—l
'*
'*   æˆøæ•ÏŠ·ƒ}ƒXƒ^b’è
'*
'*   ƒ‚ƒWƒ…[ƒ‹ID  : æˆøæ•ÏŠ·ƒ}ƒXƒ^b’è.Bas
'*
'*   ( “–ƒ‚ƒWƒ…[ƒ‹‚Í©“®¶¬‚³‚ê‚½ƒ‚ƒWƒ…[ƒ‹‚Å‚·B)
'*					/€–Ú’Ç‰ÁŠÜ‚Ş/
'*   ì¬“ú  :  2008/01/11  By Cis
'*   •ÏX“ú  :
'*---------------------------------------------------------------------------------
    Type T_æˆøæ•ÏŠ·ƒ}ƒXƒ^b’è
        Œ_–ñæ        As String
        ‘ãsŒ_–ñæ    As String
        æˆøæ        As String
        ó“ü          As String
        •i”Ô          As String
        x‹‹Œ³        As String
        x‹‹Œ³2       As String
        Œ_–ñæˆøæ    As String
        Œ_–ñHê      As String
        Œ_–ñæˆøæCD  As String
        ”[“üæ        As String
        ”[“üæHê    As String
        ”[“üæCD      As String
        ”[“üæó“ü    As String
        ”[“üæ–¼Ì    As String
        ’¼’¼‹æ•ª      As Byte
'        ì¬“ú        As Date
        ì¬Ò        As String
'        ì¬’[––      As String
        Œ”          As Long
'   €–Ú’Ç‰Á(©“®)
	—ªÌ           As String
	•i”ÔŒ”       As Long
	Œ_–ñæ”í©Ğ   As String
	Œ_–ñæ”íHê   As String
    End Type

    Public THMZ As T_æˆøæ•ÏŠ·ƒ}ƒXƒ^b’è

    Public THMZ_RDSTS    As Boolean     	    ' ê—p“Çæƒtƒ‰ƒO
    Dim adoCol        As Object
'*----------------------------------------*'
'*      Œ‹‰Ê¾¯Ä‚æ‚è\‘¢‘Ì‚Ö’l‚ğ¾¯Ä         *
'*----------------------------------------*'
Public Sub ItemsFromTHMZ(Optional DBNo As Integer)
'   ‘Ş”ğƒGƒŠƒAƒNƒŠƒA
    Call ItemsClearTHMZ	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With THMZ
            Select Case adoCol.Name
                Case "Œ_–ñæ":        If Not IsNull(adoCol.Value) Then .Œ_–ñæ       = adoCol.Value
                Case "‘ãsŒ_–ñæ":    If Not IsNull(adoCol.Value) Then .‘ãsŒ_–ñæ   = adoCol.Value
                Case "æˆøæ":        If Not IsNull(adoCol.Value) Then .æˆøæ       = adoCol.Value
                Case "ó“ü":          If Not IsNull(adoCol.Value) Then .ó“ü         = adoCol.Value
                Case "•i”Ô":          If Not IsNull(adoCol.Value) Then .•i”Ô         = adoCol.Value
                Case "x‹‹Œ³":        If Not IsNull(adoCol.Value) Then .x‹‹Œ³       = adoCol.Value
                Case "x‹‹Œ³2":       If Not IsNull(adoCol.Value) Then .x‹‹Œ³2      = adoCol.Value
                Case "Œ_–ñæˆøæ":    If Not IsNull(adoCol.Value) Then .Œ_–ñæˆøæ   = adoCol.Value
                Case "Œ_–ñHê":      If Not IsNull(adoCol.Value) Then .Œ_–ñHê     = adoCol.Value
                Case "Œ_–ñæˆøæCD":  If Not IsNull(adoCol.Value) Then .Œ_–ñæˆøæCD = adoCol.Value
                Case "”[“üæ":        If Not IsNull(adoCol.Value) Then .”[“üæ       = adoCol.Value
                Case "”[“üæHê":    If Not IsNull(adoCol.Value) Then .”[“üæHê   = adoCol.Value
                Case "”[“üæCD":      If Not IsNull(adoCol.Value) Then .”[“üæCD     = adoCol.Value
                Case "”[“üæó“ü":    If Not IsNull(adoCol.Value) Then .”[“üæó“ü   = adoCol.Value
                Case "”[“üæ–¼Ì":    If Not IsNull(adoCol.Value) Then .”[“üæ–¼Ì   = adoCol.Value
                Case "’¼’¼‹æ•ª":      If Not IsNull(adoCol.Value) Then .’¼’¼‹æ•ª     = adoCol.Value
'                Case "ì¬“ú":        If Not IsNull(adoCol.Value) Then .ì¬“ú       = adoCol.Value
                Case "ì¬Ò":        If Not IsNull(adoCol.Value) Then .ì¬Ò       = adoCol.Value
'                Case "ì¬’[––":      If Not IsNull(adoCol.Value) Then .ì¬’[––     = adoCol.Value
                Case "Œ”":         If Not IsNull(adoCol.Value) Then .Œ”        = adoCol.Value
'   €–Ú’Ç‰Á(©“®)
		Case "—ªÌ":	If Not IsNull(adoCol.Value) Then .—ªÌ	= adoCol.Value
		Case "•i”ÔŒ”":	If Not IsNull(adoCol.Value) Then .•i”ÔŒ”	= adoCol.Value
		Case "Œ_–ñæ”í©Ğ":	If Not IsNull(adoCol.Value) Then .Œ_–ñæ”í©Ğ	= adoCol.Value
		Case "Œ_–ñæ”íHê":	If Not IsNull(adoCol.Value) Then .Œ_–ñæ”íHê	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      ‘I‘ğ¸´Ø°(SELECT)Às•\‘¢‘Ì¾¯Ä    *
'*----------------------------------------*'
Public Function THMZRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    THMZRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From æˆøæ•ÏŠ·ƒ}ƒXƒ^b’è"
        RTxtSql = RTxtSql &" Where Œ_–ñæ = '" & RTrim(THMZ.Œ_–ñæ) & "'"
        RTxtSql = RTxtSql &" And   ‘ãsŒ_–ñæ = '" & RTrim(THMZ.‘ãsŒ_–ñæ) & "'"
        RTxtSql = RTxtSql &" And   æˆøæ = '" & RTrim(THMZ.æˆøæ) & "'"
        RTxtSql = RTxtSql &" And   ó“ü = '" & RTrim(THMZ.ó“ü) & "'"
        RTxtSql = RTxtSql &" And   •i”Ô = '" & RTrim(THMZ.•i”Ô) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           THMZRead = True
           Call ItemsFromTHMZ(DBNo)
	   THMZ_RDSTS = True
        else
	   THMZ_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call THMZClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      ‘I‘ğ¸´Ø°(SELECT)Às•\‘¢‘Ì¾¯Ä(Next)    *
'*----------------------------------------------*'
Public Sub THMZReadNext(Optional DBNo As Integer = 0)
    If Not THMZ_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           THMZ_RDSTS = False
        Else
           Call ItemsFromTHMZ(DBNo)
           THMZ_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     ê—pŒ‹‰Ê¾¯Ä¸Û°½Ş     *
'*-------------------------*'
Public Sub THMZClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        THMZ_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ƒ[ƒNƒGƒŠƒAƒNƒŠƒA  ( THMZ )
'*----------------------------------------------*'
Public Sub ItemsClearTHMZ()
    With THMZ
	.Œ_–ñæ = ""
	.‘ãsŒ_–ñæ = ""
	.æˆøæ = ""
	.ó“ü = ""
	.•i”Ô = ""
	.x‹‹Œ³ = ""
	.x‹‹Œ³2 = ""
	.Œ_–ñæˆøæ = ""
	.Œ_–ñHê = ""
	.Œ_–ñæˆøæCD = ""
	.”[“üæ = ""
	.”[“üæHê = ""
	.”[“üæCD = ""
	.”[“üæó“ü = ""
	.”[“üæ–¼Ì = ""
	.’¼’¼‹æ•ª = 0
'	.ì¬“ú = 0
	.ì¬Ò = ""
'	.ì¬’[–– = ""
        .Œ” = 0
'   €–Ú’Ç‰Á(©“®)
	.—ªÌ = ""
	.•i”ÔŒ” = 0
	.Œ_–ñæ”í©Ğ = ""
	.Œ_–ñæ”íHê = ""
    End With
End Sub
'*--------------------*'
'*      ’Ç‰Áˆ—       *
'*--------------------*'
Public Function THMZInsert(Optional DBNo As Integer) As Boolean
    THMZInsert = False
    With THMZ
        gSL_Select = "Insert Into æˆøæ•ÏŠ·ƒ}ƒXƒ^b’è  ( "
        gSL_Select = gSL_Select & "Œ_–ñæ,‘ãsŒ_–ñæ,æˆøæ,"
        gSL_Select = gSL_Select & "ó“ü,•i”Ô,x‹‹Œ³,"
        gSL_Select = gSL_Select & "x‹‹Œ³2,Œ_–ñæˆøæ,Œ_–ñHê,"
        gSL_Select = gSL_Select & "Œ_–ñæˆøæCD,”[“üæ,”[“üæHê,"
        gSL_Select = gSL_Select & "”[“üæCD,”[“üæó“ü,”[“üæ–¼Ì,"
        gSL_Select = gSL_Select & "’¼’¼‹æ•ª,"
        gSL_Select = gSL_Select & "ì¬“ú,ì¬Ò,ì¬’[––"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.Œ_–ñæ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.‘ãsŒ_–ñæ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.æˆøæ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ó“ü) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.•i”Ô) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.x‹‹Œ³) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.x‹‹Œ³2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.Œ_–ñæˆøæ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.Œ_–ñHê) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.Œ_–ñæˆøæCD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.”[“üæ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.”[“üæHê) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.”[“üæCD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.”[“üæó“ü) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.”[“üæ–¼Ì) & "',"		
        gSL_Select = gSL_Select & "" & .’¼’¼‹æ•ª & ","		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "'" & RTrim(.ì¬Ò) & "',"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    THMZInsert = True
End Function
'*--------------------*'
'*      XVˆ—       *
'*--------------------*'
Public Function THMZUpdate(Optional DBNo As Integer) As Boolean
    THMZUpdate = False	
    With THMZ
        gSL_Select = "Update æˆøæ•ÏŠ·ƒ}ƒXƒ^b’è  Set "
        gSL_Select = gSL_Select & "Œ_–ñæ = '" & RTrim(.Œ_–ñæ) & "',"
        gSL_Select = gSL_Select & "‘ãsŒ_–ñæ = '" & RTrim(.‘ãsŒ_–ñæ) & "',"
        gSL_Select = gSL_Select & "æˆøæ = '" & RTrim(.æˆøæ) & "',"
        gSL_Select = gSL_Select & "ó“ü = '" & RTrim(.ó“ü) & "',"
        gSL_Select = gSL_Select & "•i”Ô = '" & RTrim(.•i”Ô) & "',"
        gSL_Select = gSL_Select & "x‹‹Œ³ = '" & RTrim(.x‹‹Œ³) & "',"
        gSL_Select = gSL_Select & "x‹‹Œ³2 = '" & RTrim(.x‹‹Œ³2) & "',"
        gSL_Select = gSL_Select & "Œ_–ñæˆøæ = '" & RTrim(.Œ_–ñæˆøæ) & "',"
        gSL_Select = gSL_Select & "Œ_–ñHê = '" & RTrim(.Œ_–ñHê) & "',"
        gSL_Select = gSL_Select & "Œ_–ñæˆøæCD = '" & RTrim(.Œ_–ñæˆøæCD) & "',"
        gSL_Select = gSL_Select & "”[“üæ = '" & RTrim(.”[“üæ) & "',"
        gSL_Select = gSL_Select & "”[“üæHê = '" & RTrim(.”[“üæHê) & "',"
        gSL_Select = gSL_Select & "”[“üæCD = '" & RTrim(.”[“üæCD) & "',"
        gSL_Select = gSL_Select & "”[“üæó“ü = '" & RTrim(.”[“üæó“ü) & "',"
        gSL_Select = gSL_Select & "”[“üæ–¼Ì = '" & RTrim(.”[“üæ–¼Ì) & "',"
        gSL_Select = gSL_Select & "’¼’¼‹æ•ª = " & .’¼’¼‹æ•ª & ""
        gSL_Select = gSL_Select & " Where Œ_–ñæ = '" & RTrim(.Œ_–ñæ) & "'"
        gSL_Select = gSL_Select & "   And ‘ãsŒ_–ñæ = '" & RTrim(.‘ãsŒ_–ñæ) & "'"
        gSL_Select = gSL_Select & "   And æˆøæ = '" & RTrim(.æˆøæ) & "'"
        gSL_Select = gSL_Select & "   And ó“ü = '" & RTrim(.ó“ü) & "'"
        gSL_Select = gSL_Select & "   And •i”Ô = '" & RTrim(.•i”Ô) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    THMZUpdate = True
End Function
