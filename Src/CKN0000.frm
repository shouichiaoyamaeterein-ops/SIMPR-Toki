VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Begin VB.Form CKN0000 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   2145
   ClientLeft      =   5460
   ClientTop       =   3645
   ClientWidth     =   4245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Dummy 
      Height          =   285
      Left            =   4410
      TabIndex        =   2
      Text            =   "Dummy"
      Top             =   300
      Width           =   345
   End
   Begin VB.PictureBox PNL_Kakunin 
      BackColor       =   &H00FF0000&
      Height          =   2145
      Left            =   0
      ScaleHeight     =   2085
      ScaleWidth      =   4185
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4245
      Begin VB.PictureBox FXLabel1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1995
         Left            =   30
         ScaleHeight     =   1935
         ScaleWidth      =   4065
         TabIndex        =   1
         Top             =   30
         Width           =   4125
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   795
            Left            =   15
            Top             =   0
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   1402
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            Begin Cis3D_v60.CIS3D LBMsg 
               Height          =   720
               Left            =   45
               Top             =   30
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   1270
               BackColor       =   12648447
               ForeColor       =   16711680
               Caption         =   "   【 読 取 終 了】     再度、読取を行いますか？"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ 明朝"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   8421504
               cBoderColor2    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   1
               cPositionY      =   60
            End
         End
         Begin Cis3D_v60.CIS3D PB_OK 
            Height          =   1065
            Left            =   75
            Top             =   810
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1879
            ForeColor       =   192
            Caption         =   "続　行(Enter)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   18
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderWidth     =   9
            cAlingnment     =   7
            cPositionY      =   -170
         End
         Begin Cis3D_v60.CIS3D PB_CAN 
            Height          =   1065
            Left            =   2085
            Top             =   810
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1879
            ForeColor       =   65280
            Caption         =   "戻 る(Esc)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   18
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   14737632
            cFont3DColor2   =   8421504
            cBoderWidth     =   9
            cAlingnment     =   7
            cPositionY      =   -170
         End
      End
   End
End
Attribute VB_Name = "CKN0000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばん処理確認
'**       フォームID    :  CKN0000
'**       処理概要      :
'**
'**       作  成  日    :  2004/01/16  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim CisFrm                  As CISFormContorl
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    
    kBas_ReadOk = False
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me

    Select Case UCase(kBas_ReadMsg)
        Case "READSTART"
            With LBMsg
                .Caption = "読取を開始しますか？"
                .cAlingnment = 7
                .cPositionY = 0
                .Font.Size = 16
            End With
        Case "READEND"
            With LBMsg
                .Caption = "【 読 取 終 了】" & vbCr & "再度、読取を行いますか？"
                .cAlingnment = 6
                .cPositionY = 60
                .Font.Size = 14
            End With
        Case Else
            LBMsg.Visible = False
    End Select
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(False)
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    Call CisFrm.ExpressionForm(80, 100)
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    If Not FormAct Then Exit Function
    
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    
'   ( 続行(Enter) )
    If PB_OK.Visible Then
       If Key_Code = vbKeyReturn Then
           If IVENT = W_IVENT Then
               Call PB_OK_Click
           End If
           Key_Acc = True: Exit Function
       End If
    End If
'   ( 取消(Escape) )
    If PB_CAN.Visible Then
       If Key_Code = vbKeyEscape Then
           If IVENT = W_IVENT Then
               Call PB_CAN_Click
           End If
           Key_Acc = True: Exit Function
       End If
    End If
    
End Function
'+----------------------------+
'+    続  行  キ ー(Enter)    +
'+----------------------------+
Private Sub PB_OK_Click()
    kBas_ReadOk = True
    
    Call EndProc
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    kBas_ReadOk = False
    
    Call EndProc
End Sub
'+----------------------------+
'+       終  了  処  理        +
'+----------------------------+
Private Sub EndProc()
    Unload Me
    Set CKN0000 = Nothing
End Sub
