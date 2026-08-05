VERSION 5.00
Begin VB.Form SysMenuAbout 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ﾊﾞｰｼﾞｮﾝ情報"
   ClientHeight    =   2820
   ClientLeft      =   3675
   ClientTop       =   1515
   ClientWidth     =   5430
   ClipControls    =   0   'False
   Icon            =   "SysMenuAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1946.414
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   5099.051
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.PictureBox Picture1 
      Height          =   795
      Left            =   1560
      ScaleHeight     =   735
      ScaleWidth      =   3705
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1050
      Width           =   3765
      Begin VB.Label lb_Mem 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1050
      Picture         =   "SysMenuAbout.frx":030A
      TabIndex        =   0
      Top             =   2400
      Width           =   1310
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ｼｽﾃﾑ情報(&S)..."
      Height          =   315
      Left            =   3195
      TabIndex        =   1
      Top             =   2400
      Width           =   1310
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   240
      Picture         =   "SysMenuAbout.frx":50DC
      Stretch         =   -1  'True
      Top             =   945
      Width           =   1215
   End
   Begin VB.Label lb_Text 
      Caption         =   "Label1"
      Height          =   225
      Left            =   270
      TabIndex        =   6
      Top             =   705
      Width           =   4995
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '実線 (ふちどり)
      Index           =   1
      X1              =   56.343
      X2              =   5000.45
      Y1              =   1356.278
      Y2              =   1356.278
   End
   Begin VB.Label lb_Wsid 
      Caption         =   "ｱﾌﾟﾘｹｰｼｮﾝの説明"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   270
      TabIndex        =   2
      Top             =   450
      Width           =   4935
   End
   Begin VB.Label lblTitle 
      Caption         =   "ｱﾌﾟﾘｹｰｼｮﾝ ﾀｲﾄﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   270
      TabIndex        =   4
      Top             =   180
      Width           =   2625
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   70.429
      X2              =   5000.45
      Y1              =   1366.631
      Y2              =   1366.631
   End
   Begin VB.Label lblVersion 
      Caption         =   "ﾊﾞｰｼﾞｮﾝ"
      Height          =   165
      Left            =   2910
      TabIndex        =   5
      Top             =   180
      Width           =   2085
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Chukyo Infomation System Co.,LTD."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Index           =   0
      Left            =   495
      TabIndex        =   3
      Top             =   2055
      Width           =   4470
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Chukyo Infomation System Co.,LTD."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   525
      TabIndex        =   7
      Top             =   2070
      Width           =   4470
   End
End
Attribute VB_Name = "SysMenuAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim CisFrm          As CISFormContorl

' ﾚｼﾞｽﾄﾘ ｷｰ ｾｷｭﾘﾃｨ ｵﾌﾟｼｮﾝ...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ﾚｼﾞｽﾄﾘ ｷｰ ROOT 型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode Null 文字で終わる文字列
Const REG_DWORD = 4                      ' 32 ﾋﾞｯﾄ数値

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Type MEMORYSTATUS
  dwLength As Long        '構造体のｻｲｽﾞ
  dwMemoryLoad As Long    '使用中のﾒﾓﾘ(%)
  dwTotalPhys As Long     '物理ﾒﾓﾘ(ﾊﾞｲﾄ)
  dwAvailPhys As Long     '未使用ﾒﾓﾘ(ﾊﾞｲﾄ)
  dwTotalPageFile As Long 'ﾍﾟｰｼﾞﾝｸﾞﾌｧｲﾙのｻｲｽﾞ(ﾊﾞｲﾄ)
  dwAvailPageFile As Long '使用可能ﾍﾟｰｼﾞﾝｸﾞﾌｧｲﾙのｻｲｽﾞ(ﾊﾞｲﾄ)
  dwTotalVirtual As Long  'ﾕｰｻﾞｰが使える仮想ﾒﾓﾘ(ﾊﾞｲﾄ）
  dwAvailVirtual As Long  'ﾕｰｻﾞｰが使える未使用ﾒﾓﾘ(ﾊﾞｲﾄ)
End Type
Dim MS      As MEMORYSTATUS
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub CmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Call GlobalMemoryStatus(MS)

  lb_Mem = "Windows が使用できる物理ﾒﾓﾘ ： " & Format(MS.dwTotalPhys / 1024, "#,###") & " KB"
'  If ChkTask("SQLAgent.exe") Then
'     lb_Text = "SQLｽｹｼﾞｭｰﾙ起動中"
'  End If
End Sub

Private Sub Form_Load()
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Me.Caption = "Version Information"

    With CisFun
        .INI_Section = "System"
        .INI_Key = "Ver"
        .INI_Default = ""
        .INI_FileID = gIniExe & gIniName
        
        .GetIni
        
        If Trim(.INI_String) = "" Then
            lblVersion.Caption = "Version  " & App.Major & "." & App.Minor & "." & App.Revision
        Else
            lblVersion.Caption = "Version  " & Mid(.INI_String, 1, 1) & "." & Mid(.INI_String, 3, 1) & "." & Mid(.INI_String, 4, 1)
        End If
    End With
    
    lblTitle.Caption = App.Title
    lb_Wsid = "ｺﾝﾋﾟｭｰﾀ名 ： " & CisFun.WSID & " / DB名 ： " & CisDB.DBName
    lb_Text = "  Menu Text    ： " & gFT_Text
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ﾚｼﾞｽﾄﾘからｼｽﾃﾑ情報ﾌﾟﾛｸﾞﾗﾑのﾊﾟｽ\名前を取得します...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' ﾚｼﾞｽﾄﾘからｼｽﾃﾑ情報ﾌﾟﾛｸﾞﾗﾑのﾊﾟｽ名のみを取得します...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' 既に存在するはずの 32 ﾋﾞｯﾄ ﾊﾞｰｼﾞｮﾝのﾌｧｲﾙを確認します。
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ｴﾗｰ - ﾌｧｲﾙが見つかりません...
        Else
            GoTo SysInfoErr
        End If
    ' ｴﾗｰ - ﾚｼﾞｽﾄﾘ ｴﾝﾄﾘが見つかりません...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "現時点ではｼｽﾃﾑ情報を使用できません", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' ﾙｰﾌﾟ ｶｳﾝﾀ
    Dim rc As Long                                          ' 戻り値のｺｰﾄﾞ
    Dim hKey As Long                                        ' ｵｰﾌﾟﾝしたﾚｼﾞｽﾄﾘ ｷｰのﾊﾝﾄﾞﾙ
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ﾚｼﾞｽﾄﾘ ｷｰのﾃﾞｰﾀ型
    Dim tmpVal As String                                    ' ﾚｼﾞｽﾄﾘ ｷｰ値の一時保存領域
    Dim KeyValSize As Long                                  ' ﾚｼﾞｽﾄﾘ ｷｰ変数のｻｲｽﾞ
    '------------------------------------------------------------
    ' ﾙｰﾄ ｷｰ {HKEY_LOCAL_MACHINE...} にあるﾚｼﾞｽﾄﾘ ｷｰを開きます。
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ﾚｼﾞｽﾄﾘ ｷｰを開く
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ﾊﾝﾄﾞﾙ ｴﾗｰ...
    
    tmpVal = String$(1024, 0)                             ' 変数領域の割り当て
    KeyValSize = 1024                                       ' 変数のｻｲｽﾞを記憶
    
    '------------------------------------------------------------
    ' ﾚｼﾞｽﾄﾘ ｷｰ値を取得します...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ｷｰ値の取得/作成
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ﾊﾝﾄﾞﾙ ｴﾗｰ
    
    tmpVal = Left(tmpVal, InStr(tmpVal, Chr(0)) - 1)
    '------------------------------------------------------------
    ' 変換のために、ｷｰ値の型を調べます...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' ﾃﾞｰﾀ型検索...
    Case REG_SZ                                             ' String ﾚｼﾞｽﾄﾘ ｷｰ ﾃﾞｰﾀ型
        KeyVal = tmpVal                                     ' String 値をｺﾋﾟｰ
    Case REG_DWORD                                          ' Double Word ﾚｼﾞｽﾄﾘ ｷｰ ﾃﾞｰﾀ型
        For i = Len(tmpVal) To 1 Step -1                    ' 各ﾋﾞｯﾄの変換
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Char ごとに値を作成
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Double Word を String に変換
    End Select
    
    GetKeyValue = True                                      ' 正常終了
    rc = RegCloseKey(hKey)                                  ' ﾚｼﾞｽﾄﾘ ｷｰをｸﾛｰｽﾞ
    Exit Function                                           ' 終了
    
GetKeyError:      ' ｴﾗｰ発生後の後始末...
    KeyVal = ""                                             ' 戻り値の値を空文字列に設定
    GetKeyValue = False                                     ' 異常終了
    rc = RegCloseKey(hKey)                                  ' ﾚｼﾞｽﾄﾘ ｷｰをｸﾛｰｽﾞ
End Function
