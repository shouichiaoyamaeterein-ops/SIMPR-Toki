VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "SQL Server の接続"
   ClientHeight    =   2565
   ClientLeft      =   3975
   ClientTop       =   3405
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515.486
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   4478.772
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDBName 
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   780
      Width           =   2325
   End
   Begin VB.TextBox txtServer 
      Height          =   345
      Left            =   2025
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2040
      TabIndex        =   2
      Top             =   1170
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "接続(&O)"
      Default         =   -1  'True
      Height          =   300
      Left            =   1905
      TabIndex        =   6
      Top             =   2145
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ(&C)"
      Height          =   300
      Left            =   3510
      TabIndex        =   7
      Top             =   2145
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1560
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "ﾃﾞｰﾀﾍﾞｰｽ名  ："
      Height          =   270
      Index           =   3
      Left            =   720
      TabIndex        =   9
      Top             =   855
      Width           =   1290
   End
   Begin VB.Label lblLabels 
      Caption         =   "サーバー名　 ："
      Height          =   270
      Index           =   2
      Left            =   720
      TabIndex        =   8
      Top             =   210
      Width           =   1230
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   4478.772
      Y1              =   381.087
      Y2              =   381.087
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   -28.168
      X2              =   4450.604
      Y1              =   372.225
      Y2              =   372.225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   28.168
      X2              =   4506.94
      Y1              =   1214.161
      Y2              =   1214.161
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   4478.772
      Y1              =   1205.299
      Y2              =   1205.299
   End
   Begin VB.Label lblLabels 
      Caption         =   "ログイン名   ："
      Height          =   270
      Index           =   0
      Left            =   795
      TabIndex        =   4
      Top             =   1245
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "パスワード   ："
      Height          =   270
      Index           =   1
      Left            =   780
      TabIndex        =   5
      Top             =   1635
      Width           =   1170
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**       変  更  日    :   2005/09/27  By CIS ﾛｸﾞｲﾝ情報入力後、ﾊﾟｽﾜｰﾄﾞ更新が行われていなかった。
    Dim CisFrm          As CISFormContorl

Dim InputObj          As Boolean

Private Sub cmdCancel_Click()
    'ﾛｸﾞｵﾝに失敗したことを示すために、
    'ｸﾞﾛｰﾊﾞﾙ変数を Flase に設定します。
    gLoginSucceeded = False
    Me.Hide
End Sub

Private Sub CmdOK_Click()
    If Trim(txtServer) = "" Then
        MsgBox "サーバー名を入力して下さい。!", vbExclamation, "ﾛｸﾞｵﾝ"
        txtServer.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim(txtDBName) = "" Then
        MsgBox "ﾃﾞｰﾀﾍﾞｰｽ名を入力して下さい。!", vbExclamation, "ﾛｸﾞｵﾝ"
        txtDBName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim(txtUserName) = "" Then
        MsgBox "ログイン名を入力して下さい。!", vbExclamation, "ﾛｸﾞｵﾝ"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    With CisDB
        .User = txtUserName
        .PASS = txtPassword
        .DBName = txtDBName
        .Server = txtServer
    End With
    
'   ( INI File 書込 )
    With CisFun
        .INI_FileID = gIniExe & gIniName
        .INI_Section = "DataBase"
        
        .INI_Key = "Server"
        .INI_String = txtServer
        .PutIni
    
        .INI_Key = "Name"
        .INI_String = txtDBName
        .PutIni
    
        .INI_Key = "User"
        .INI_String = txtUserName
        .PutIni
    
''       ( パスワード変換 )
'        sSapWK = RTrim(txtPassword) & Space(15)
'        Call PASS(sSAP15, sSapWK)
        
        .INI_Key = "PASS"
'        .INI_String = sSapWK
        .INI_String = txtPassword
'        .PutIni
    
        '+------------------------------------------------------- 2005/09/27 Insert
        gStr = CisDB.PassWord(RTrim(txtPassword))
        Call SaveSetting(gRegKey, "Environment", "DrowsSap", gStr)
        '+------------------------------------------------------- 2005/09/27 Insert End
    
    End With
    
    Call CisFrm.MousePT(11)
    Me.Move -10000
    
    If Not gDBLogon Then
        If Not IniGet Then
            With CisFun
                .MB_Lines = 6
                .MB_MSG(1) = "   データベースに接続出来ません！！"
                .MB_MSG(5) = "  設定を見直して再度実行して下さい　　"
                .MB_Title = "ﾃﾞｰﾀﾍﾞｰｽ接続"
                .MB_Button = Error
                .MBOX
            End With
            
            Call CisFrm.MousePT(1)
            Call CisFrm.Frm_Center(False)
            Exit Sub
        End If
    End If
    
    gLoginSucceeded = True
    Me.Hide
End Sub

Private Sub Form_Load()
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    
    With CisDB
        txtServer = .Server
        If Trim(txtServer) = "" Then
           txtServer = CisFun.WSID
        End If
        txtUserName = .User
        txtDBName = .DBName
        If txtServer <> "" Then txtPassword = .PASS
    End With

'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(False)
    
End Sub
Private Sub Form_Activate()
    If InputObj Then
        txtPassword.SetFocus
    End If
End Sub
