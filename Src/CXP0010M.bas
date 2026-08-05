Attribute VB_Name = "CXP0010M"
Option Explicit

Public Pass_Temp       As String
Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Public Const Pass_FName     As String = "vbp0311.tmp"

Dim P_Fno               As Integer

Global In_Pass          As String
Global In_Tanto         As String

Global ck_Pass          As String
Global ck_Tanto         As String
Dim pStr                As String

Global P_Area(2)        As String
'+------------------------------+
'+   Tempフォルダを求める        +
'+------------------------------+
Public Sub GetTemp()
    Dim strEnvironmentSrc    As String
    Dim strEnvironmentDest   As String * 256
    Dim lngEnvironmentLenght As Long
' 展開する環境変数文字列を設定
    strEnvironmentSrc = "%windir%"
    strEnvironmentSrc = "%temp%"
' 環境変数文字列を展開
    lngEnvironmentLenght = ExpandEnvironmentStrings(strEnvironmentSrc, strEnvironmentDest, Len(strEnvironmentDest))
' 環境変数文字列を表示
    Pass_Temp = Left(strEnvironmentDest, InStr(strEnvironmentDest, vbNullChar) - 1)
    If Trim(Pass_Temp) = "" Then Exit Sub
    If Right(Pass_Temp, 1) <> "\" Then
       Pass_Temp = Pass_Temp & "\"
    End If

    In_Pass = "": In_Tanto = ""
    If Dir(Pass_Temp & Pass_FName) = "" Then Exit Sub
        

    P_Fno = FreeFile
    
    Open Pass_Temp & Pass_FName For Input As #P_Fno
    
    Line Input #P_Fno, pStr
    In_Pass = CisDB.PassWord(pStr)
    
    Line Input #P_Fno, pStr
    In_Tanto = CisDB.PassWord(pStr)
    
    Close #P_Fno
    


End Sub


