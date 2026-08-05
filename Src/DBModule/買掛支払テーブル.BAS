Attribute VB_Name = "”ƒŠ|x•¥ƒe[ƒuƒ‹"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) d—l
'*
'*   ”ƒŠ|x•¥ƒe[ƒuƒ‹
'*
'*   ƒ‚ƒWƒ…[ƒ‹ID  : ”ƒŠ|x•¥ƒe[ƒuƒ‹.Bas
'*
'*   ( “–ƒ‚ƒWƒ…[ƒ‹‚Í©“®¶¬‚³‚ê‚½ƒ‚ƒWƒ…[ƒ‹‚Å‚·B)
'*					/€–Ú’Ç‰ÁŠÜ‚Ş/
'*   ì¬“ú  :  2004/12/08  By Cis
'*   •ÏX“ú  :
'*---------------------------------------------------------------------------------
    Type T_”ƒŠ|x•¥ƒe[ƒuƒ‹
        ÀÑ”NŒ          As String
        Œ_–ñæ            As String
        ‘OŒŒJ‰z          As Long
        d“ü‹àŠz          As Long
        d“üÁ”ïÅ        As Long
        ‘ŠE‹àŠz          As Long
        ‘ŠEÁ”ïÅ        As Long
        ¿‹‹àŠz          As Long
        ¿‹‘NO          As String
        x•¥‹àŠz          As Long
        ŒJ‰z‹àŠz          As Long
        x•¥‹æ•ª1         As String
        x•¥—¦1           As Byte
        x•¥“ú1           As String
        x•¥è”—¿1       As Long
        x•¥‹àŠz1         As Long
        x•¥‹æ•ª2         As String
        x•¥—¦2           As Byte
        x•¥“ú2           As String
        x•¥è”—¿2       As Long
        x•¥‹àŠz2         As Long
        ¿‹Œ”          As Long
        ’÷“ú              As Byte
'        ì¬“ú            As Date
'        ì¬’[––          As String
        ì¬Ò            As String
'        XV“ú            As Date
'        XV’[––          As String
        XVÒ            As String
        Œ”          As Long
'   €–Ú’Ç‰Á(©“®)
	Œ_–ñæ–¼      As String
	x•¥‹æ•ª–¼1   As String
	x•¥‹æ•ª–¼2   As String
    End Type

    Public KSH As T_”ƒŠ|x•¥ƒe[ƒuƒ‹

    Public KSH_RDSTS    As Boolean     	    ' ê—p“Çæƒtƒ‰ƒO
    Dim adoCol        As Object
'*----------------------------------------*'
'*      Œ‹‰Ê¾¯Ä‚æ‚è\‘¢‘Ì‚Ö’l‚ğ¾¯Ä         *
'*----------------------------------------*'
Public Sub ItemsFromKSH(Optional DBNo As Integer)
'   ‘Ş”ğƒGƒŠƒAƒNƒŠƒA
    Call ItemsClearKSH	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KSH
            Select Case adoCol.Name
                Case "ÀÑ”NŒ":          If Not IsNull(adoCol.Value) Then .ÀÑ”NŒ         = adoCol.Value
                Case "Œ_–ñæ":            If Not IsNull(adoCol.Value) Then .Œ_–ñæ           = adoCol.Value
                Case "‘OŒŒJ‰z":          If Not IsNull(adoCol.Value) Then .‘OŒŒJ‰z         = adoCol.Value
                Case "d“ü‹àŠz":          If Not IsNull(adoCol.Value) Then .d“ü‹àŠz         = adoCol.Value
                Case "d“üÁ”ïÅ":        If Not IsNull(adoCol.Value) Then .d“üÁ”ïÅ       = adoCol.Value
                Case "‘ŠE‹àŠz":          If Not IsNull(adoCol.Value) Then .‘ŠE‹àŠz         = adoCol.Value
                Case "‘ŠEÁ”ïÅ":        If Not IsNull(adoCol.Value) Then .‘ŠEÁ”ïÅ       = adoCol.Value
                Case "¿‹‹àŠz":          If Not IsNull(adoCol.Value) Then .¿‹‹àŠz         = adoCol.Value
                Case "¿‹‘NO":          If Not IsNull(adoCol.Value) Then .¿‹‘NO         = adoCol.Value
                Case "x•¥‹àŠz":          If Not IsNull(adoCol.Value) Then .x•¥‹àŠz         = adoCol.Value
                Case "ŒJ‰z‹àŠz":          If Not IsNull(adoCol.Value) Then .ŒJ‰z‹àŠz         = adoCol.Value
                Case "x•¥‹æ•ª1":         If Not IsNull(adoCol.Value) Then .x•¥‹æ•ª1        = adoCol.Value
                Case "x•¥—¦1":           If Not IsNull(adoCol.Value) Then .x•¥—¦1          = adoCol.Value
                Case "x•¥“ú1":           If Not IsNull(adoCol.Value) Then .x•¥“ú1          = adoCol.Value
                Case "x•¥è”—¿1":       If Not IsNull(adoCol.Value) Then .x•¥è”—¿1      = adoCol.Value
                Case "x•¥‹àŠz1":         If Not IsNull(adoCol.Value) Then .x•¥‹àŠz1        = adoCol.Value
                Case "x•¥‹æ•ª2":         If Not IsNull(adoCol.Value) Then .x•¥‹æ•ª2        = adoCol.Value
                Case "x•¥—¦2":           If Not IsNull(adoCol.Value) Then .x•¥—¦2          = adoCol.Value
                Case "x•¥“ú2":           If Not IsNull(adoCol.Value) Then .x•¥“ú2          = adoCol.Value
                Case "x•¥è”—¿2":       If Not IsNull(adoCol.Value) Then .x•¥è”—¿2      = adoCol.Value
                Case "x•¥‹àŠz2":         If Not IsNull(adoCol.Value) Then .x•¥‹àŠz2        = adoCol.Value
                Case "¿‹Œ”":          If Not IsNull(adoCol.Value) Then .¿‹Œ”         = adoCol.Value
                Case "’÷“ú":              If Not IsNull(adoCol.Value) Then .’÷“ú             = adoCol.Value
'                Case "ì¬“ú":            If Not IsNull(adoCol.Value) Then .ì¬“ú           = adoCol.Value
'                Case "ì¬’[––":          If Not IsNull(adoCol.Value) Then .ì¬’[––         = adoCol.Value
                Case "ì¬Ò":            If Not IsNull(adoCol.Value) Then .ì¬Ò           = adoCol.Value
'                Case "XV“ú":            If Not IsNull(adoCol.Value) Then .XV“ú           = adoCol.Value
'                Case "XV’[––":          If Not IsNull(adoCol.Value) Then .XV’[––         = adoCol.Value
                Case "XVÒ":            If Not IsNull(adoCol.Value) Then .XVÒ           = adoCol.Value
                Case "Œ”":         If Not IsNull(adoCol.Value) Then .Œ”        = adoCol.Value
'   €–Ú’Ç‰Á(©“®)
		Case "Œ_–ñæ–¼":	If Not IsNull(adoCol.Value) Then .Œ_–ñæ–¼	= adoCol.Value
		Case "x•¥‹æ•ª–¼1":	If Not IsNull(adoCol.Value) Then .x•¥‹æ•ª–¼1	= adoCol.Value
		Case "x•¥‹æ•ª–¼2":	If Not IsNull(adoCol.Value) Then .x•¥‹æ•ª–¼2	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      ‘I‘ğ¸´Ø°(SELECT)Às•\‘¢‘Ì¾¯Ä    *
'*----------------------------------------*'
Public Function KSHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KSHRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From ”ƒŠ|x•¥ƒe[ƒuƒ‹"
        RTxtSql = RTxtSql &" Where ÀÑ”NŒ = '" & RTrim(KSH.ÀÑ”NŒ) & "'"
        RTxtSql = RTxtSql &" And   Œ_–ñæ = '" & RTrim(KSH.Œ_–ñæ) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KSHRead = True
           Call ItemsFromKSH(DBNo)
	   KSH_RDSTS = True
        else
	   KSH_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KSHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      ‘I‘ğ¸´Ø°(SELECT)Às•\‘¢‘Ì¾¯Ä(Next)    *
'*----------------------------------------------*'
Public Sub KSHReadNext(Optional DBNo As Integer = 0)
    If Not KSH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KSH_RDSTS = False
        Else
           Call ItemsFromKSH(DBNo)
           KSH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     ê—pŒ‹‰Ê¾¯Ä¸Û°½Ş     *
'*-------------------------*'
Public Sub KSHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KSH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ƒ[ƒNƒGƒŠƒAƒNƒŠƒA  ( KSH )
'*----------------------------------------------*'
Public Sub ItemsClearKSH()
    With KSH
	.ÀÑ”NŒ = ""
	.Œ_–ñæ = ""
	.‘OŒŒJ‰z = 0
	.d“ü‹àŠz = 0
	.d“üÁ”ïÅ = 0
	.‘ŠE‹àŠz = 0
	.‘ŠEÁ”ïÅ = 0
	.¿‹‹àŠz = 0
	.¿‹‘NO = ""
	.x•¥‹àŠz = 0
	.ŒJ‰z‹àŠz = 0
	.x•¥‹æ•ª1 = ""
	.x•¥—¦1 = 0
	.x•¥“ú1 = ""
	.x•¥è”—¿1 = 0
	.x•¥‹àŠz1 = 0
	.x•¥‹æ•ª2 = ""
	.x•¥—¦2 = 0
	.x•¥“ú2 = ""
	.x•¥è”—¿2 = 0
	.x•¥‹àŠz2 = 0
	.¿‹Œ” = 0
	.’÷“ú = 0
'	.ì¬“ú = 0
'	.ì¬’[–– = ""
	.ì¬Ò = ""
'	.XV“ú = 0
'	.XV’[–– = ""
	.XVÒ = ""
        .Œ” = 0
'   €–Ú’Ç‰Á(©“®)
	.Œ_–ñæ–¼ = ""
	.x•¥‹æ•ª–¼1 = ""
	.x•¥‹æ•ª–¼2 = ""
    End With
End Sub
'*--------------------*'
'*      ’Ç‰Áˆ—       *
'*--------------------*'
Public Function KSHInsert(Optional DBNo As Integer) As Boolean
    KSHInsert = True
    With KSH
        gSL_Select = "Insert Into ”ƒŠ|x•¥ƒe[ƒuƒ‹  ( "
        gSL_Select = gSL_Select & "ÀÑ”NŒ,Œ_–ñæ,‘OŒŒJ‰z,"
        gSL_Select = gSL_Select & "d“ü‹àŠz,d“üÁ”ïÅ,‘ŠE‹àŠz,"
        gSL_Select = gSL_Select & "‘ŠEÁ”ïÅ,¿‹‹àŠz,¿‹‘NO,"
        gSL_Select = gSL_Select & "x•¥‹àŠz,ŒJ‰z‹àŠz,x•¥‹æ•ª1,"
        gSL_Select = gSL_Select & "x•¥—¦1,x•¥“ú1,x•¥è”—¿1,"
        gSL_Select = gSL_Select & "x•¥‹àŠz1,x•¥‹æ•ª2,x•¥—¦2,"
        gSL_Select = gSL_Select & "x•¥“ú2,x•¥è”—¿2,x•¥‹àŠz2,"
        gSL_Select = gSL_Select & "¿‹Œ”,’÷“ú,"
        gSL_Select = gSL_Select & "ì¬“ú,ì¬’[––,ì¬Ò"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.ÀÑ”NŒ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.Œ_–ñæ) & "',"		
        gSL_Select = gSL_Select & "" & .‘OŒŒJ‰z & ","		
        gSL_Select = gSL_Select & "" & .d“ü‹àŠz & ","		
        gSL_Select = gSL_Select & "" & .d“üÁ”ïÅ & ","		
        gSL_Select = gSL_Select & "" & .‘ŠE‹àŠz & ","		
        gSL_Select = gSL_Select & "" & .‘ŠEÁ”ïÅ & ","		
        gSL_Select = gSL_Select & "" & .¿‹‹àŠz & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.¿‹‘NO) & "',"		
        gSL_Select = gSL_Select & "" & .x•¥‹àŠz & ","		
        gSL_Select = gSL_Select & "" & .ŒJ‰z‹àŠz & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.x•¥‹æ•ª1) & "',"		
        gSL_Select = gSL_Select & "" & .x•¥—¦1 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.x•¥“ú1) & "',"		
        gSL_Select = gSL_Select & "" & .x•¥è”—¿1 & ","		
        gSL_Select = gSL_Select & "" & .x•¥‹àŠz1 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.x•¥‹æ•ª2) & "',"		
        gSL_Select = gSL_Select & "" & .x•¥—¦2 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.x•¥“ú2) & "',"		
        gSL_Select = gSL_Select & "" & .x•¥è”—¿2 & ","		
        gSL_Select = gSL_Select & "" & .x•¥‹àŠz2 & ","		
        gSL_Select = gSL_Select & "" & .¿‹Œ” & ","		
        gSL_Select = gSL_Select & "" & .’÷“ú & ","		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"		
        gSL_Select = gSL_Select & "'" & RTrim(.ì¬Ò) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KSHInsert = True
End Function
'*--------------------*'
'*      XVˆ—       *
'*--------------------*'
Public Function KSHUpdate(Optional DBNo As Integer) As Boolean
    KSHUpdate = False	
    With KSH
        gSL_Select = "Update ”ƒŠ|x•¥ƒe[ƒuƒ‹  Set "
        gSL_Select = gSL_Select & "ÀÑ”NŒ = '" & RTrim(.ÀÑ”NŒ) & "',"
        gSL_Select = gSL_Select & "Œ_–ñæ = '" & RTrim(.Œ_–ñæ) & "',"
        gSL_Select = gSL_Select & "‘OŒŒJ‰z = " & .‘OŒŒJ‰z & ","
        gSL_Select = gSL_Select & "d“ü‹àŠz = " & .d“ü‹àŠz & ","
        gSL_Select = gSL_Select & "d“üÁ”ïÅ = " & .d“üÁ”ïÅ & ","
        gSL_Select = gSL_Select & "‘ŠE‹àŠz = " & .‘ŠE‹àŠz & ","
        gSL_Select = gSL_Select & "‘ŠEÁ”ïÅ = " & .‘ŠEÁ”ïÅ & ","
        gSL_Select = gSL_Select & "¿‹‹àŠz = " & .¿‹‹àŠz & ","
        gSL_Select = gSL_Select & "¿‹‘NO = '" & RTrim(.¿‹‘NO) & "',"
        gSL_Select = gSL_Select & "x•¥‹àŠz = " & .x•¥‹àŠz & ","
        gSL_Select = gSL_Select & "ŒJ‰z‹àŠz = " & .ŒJ‰z‹àŠz & ","
        gSL_Select = gSL_Select & "x•¥‹æ•ª1 = '" & RTrim(.x•¥‹æ•ª1) & "',"
        gSL_Select = gSL_Select & "x•¥—¦1 = " & .x•¥—¦1 & ","
        gSL_Select = gSL_Select & "x•¥“ú1 = '" & RTrim(.x•¥“ú1) & "',"
        gSL_Select = gSL_Select & "x•¥è”—¿1 = " & .x•¥è”—¿1 & ","
        gSL_Select = gSL_Select & "x•¥‹àŠz1 = " & .x•¥‹àŠz1 & ","
        gSL_Select = gSL_Select & "x•¥‹æ•ª2 = '" & RTrim(.x•¥‹æ•ª2) & "',"
        gSL_Select = gSL_Select & "x•¥—¦2 = " & .x•¥—¦2 & ","
        gSL_Select = gSL_Select & "x•¥“ú2 = '" & RTrim(.x•¥“ú2) & "',"
        gSL_Select = gSL_Select & "x•¥è”—¿2 = " & .x•¥è”—¿2 & ","
        gSL_Select = gSL_Select & "x•¥‹àŠz2 = " & .x•¥‹àŠz2 & ","
        gSL_Select = gSL_Select & "¿‹Œ” = " & .¿‹Œ” & ","
        gSL_Select = gSL_Select & "’÷“ú = " & .’÷“ú & ","
        gSL_Select = gSL_Select & "XV“ú = GETDATE(),"
        gSL_Select = gSL_Select & "XV’[–– = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "XVÒ = '" & RTrim(.XVÒ) & "'"
        gSL_Select = gSL_Select & " Where ÀÑ”NŒ = '" & RTrim(.ÀÑ”NŒ) & "'"
        gSL_Select = gSL_Select & "   And Œ_–ñæ = '" & RTrim(.Œ_–ñæ) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KSHUpdate = True
End Function
