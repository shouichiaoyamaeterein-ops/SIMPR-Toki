VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~2.OCX"
Object = "{0D300FC0-B2EA-11D1-8D3B-444553540000}#1.30#0"; "Qrmaker.ocx"
Begin VB.Form CKN1240 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "印刷中"
   ClientHeight    =   1755
   ClientLeft      =   4845
   ClientTop       =   4575
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5640
   Begin VB.CommandButton PB_Can 
      Caption         =   "印刷中止"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1980
      TabIndex        =   2
      Top             =   1260
      Width           =   1635
   End
   Begin QRMAKERLib.QRmaker QRmaker1 
      Height          =   585
      Left            =   30
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   615
      _Version        =   65566
      _ExtentX        =   1085
      _ExtentY        =   1032
      _StockProps     =   1
      BackColor       =   16761024
      Picture         =   "CKN1240.frx":0000
   End
   Begin Cis3D_v60.CIS3D Prt_Cnt 
      Height          =   330
      Left            =   1725
      Top             =   810
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   582
      BackColor       =   16777152
      ForeColor       =   0
      Caption         =   "Z,ZZZ / Z,ZZZ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   8421504
      cFont3DColor2   =   -2147483643
      cFont3D         =   2
      cAlingnment     =   7
   End
   Begin CisMsg_v60.CISMSG Prt_Msg 
      Height          =   1800
      Left            =   -15
      TabIndex        =   0
      Top             =   -45
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   3175
      BackColor       =   -2147483633
      ForeColor       =   16777215
      Caption         =   "【  印 刷 中  】"
      BackStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   0
      cFont3DColor2   =   8421504
      cAlingnment     =   7
      cPositionY      =   -400
      gGC1_Red        =   180
      gGC2_Green      =   200
      gGC3_Bule       =   100
      gGradation      =   -1  'True
   End
End
Attribute VB_Name = "CKN1240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばん 発行
'**       フォームID    :  CKN1240
'**       処理概要      :
'**
'**       作  成  日    :  2004/09/02  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim mLoop           As Long

    Dim mPrintName      As String
    
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld

'   #----------------#
'   #    初期設定    #
'   #----------------#
    Prt_Msg = pref_Msg
    Prt_Cnt = ""
    FormAct = False

'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(False)
    
    DoEvents
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
        
    Call CisFrm.MousePT(11)
    
    Set mQRMeker = QRmaker1
    pref_Cansel_Flg = False
    
    For mLoop = 1 To 200
        DoEvents
    Next mLoop
    
    Set pref_Cnt = Prt_Cnt

    ' 発行用設定項目
    With CisFun
        .INI_Section = "かんばん印刷先"
        .INI_FileID = gIniExe & gIniName
        ' 印刷先
        .INI_Key = CisFun.WSID
        .INI_Default = ""
        .GetIni
        gStr = RTrim(.INI_String)
        If Trim(gStr) = "" Then
           For gCnt1 = 0 To Printers.Count - 1
               If Printers(gCnt1).DeviceName Like "TEC*" Then
                  gStr = Printers(gCnt1).DeviceName
                  Exit For
               End If
           Next gCnt1
           If Trim(gStr) <> "" Then
              .INI_String = gStr
              .PutIni
           End If
        End If
        mPrintName = gStr
    End With
'+++++++++++++++++++++++++++++++'
'+  プリンタ切替
'+++++++++++++++++++++++++++++++'
    Call CisFrm.PrintSet(mPrintName)
' 印刷処理
    pref_Rtn = False
        
    If CoPrintN1240 Then
       pref_Rtn = True
    End If
    
    Call CisFrm.MousePT(1)

    Call CisFrm.PrintSet("")

' 処理終了
    Call ProcEnd
    
End Sub
'+----------------------+
'+      終了処理        +
'+----------------------+
Private Sub ProcEnd()
    Me.Hide
End Sub
'+----------------------------+
'+    印刷取消キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    pref_Cansel_Flg = True
End Sub

