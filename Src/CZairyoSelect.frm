VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form CZairyoSelect 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料選択画面"
   ClientHeight    =   3375
   ClientLeft      =   5295
   ClientTop       =   4890
   ClientWidth     =   10275
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   10275
   Begin VB.PictureBox Picture 
      Height          =   2610
      Index           =   1
      Left            =   30
      ScaleHeight     =   2550
      ScaleWidth      =   9585
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   9645
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CZairyoSelect.frx":0000
         Height          =   2520
         Left            =   30
         OleObjectBlob   =   "CZairyoSelect.frx":0015
         TabIndex        =   4
         Top             =   30
         Width           =   9525
      End
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H00404040&
      Height          =   660
      Index           =   2
      Left            =   8055
      ScaleHeight     =   600
      ScaleWidth      =   1560
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   1620
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   555
         Left            =   780
         Top             =   30
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnType         =   2
         BtnCaption      =   "入力"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   128
         BtnShadowColor  =   16777215
         PFCaption       =   "(Enter)"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   16777215
         PFShadowColor   =   0
         PFBackColor     =   128
      End
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   555
         Left            =   30
         Top             =   30
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnType         =   3
         BtnCaption      =   "取消"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   32768
         BtnShadowColor  =   16777215
         PFCaption       =   "( Esc )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   49152
      End
   End
   Begin VB.PictureBox Picture 
      Height          =   630
      Index           =   0
      Left            =   30
      ScaleHeight     =   570
      ScaleWidth      =   6165
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   6225
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   465
         Left            =   4215
         Top             =   60
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "背番号"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D H1lb_Seban 
            Height          =   360
            Left            =   810
            Top             =   60
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXX"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   465
         Left            =   60
         Top             =   60
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "取引先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   870
            Top             =   60
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   -150
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   6720
      Top             =   690
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CZairyoSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  受給/発注品番選択画面
'**       フォームID    :  CZairyoSelect
'**       処理概要      :
'**
'**       作  成  日    :  2008/02/29  By CIS
'**       変  更  日    :   2014/02/04  By CIS WIN7対応
'**       変  更  日    :
'**
'**********************************************************'
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim MeCap           As String
    Dim mSL_Select      As String       ' Select 文
 
    Dim wObj            As Object       ' ｵﾌﾞｼﾞｪｸﾄﾜｰｸｴﾘｱ
    Dim CisFrm          As CISFormContorl
 
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    Dummy.Left = -1000
    Dummy.Enabled = False
'
'   #--------------#
'   # 画面表示位置  #
'   #--------------#
    If RV_Left <> 0 Or RV_Top <> 0 Then
       Me.Move RV_Left, RV_Top
    Else
       Call CisFrm.Frm_Center(False, x1024y768)
    End If

End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    RV_Rtn = False
    
    MeCap = Me.Caption
    
    Adodc1.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
 
    mSL_Select = ""
    mSL_Select = mSL_Select & " SELECT ZT.材料管理番号, ZT.受給品番, ZT.発注品番,"
    mSL_Select = mSL_Select & "        ZT.材質, ZM.板厚, ZM.幅, ZM.長さ,"
    mSL_Select = mSL_Select & "        ZT.材質 + ' ' +"
    mSL_Select = mSL_Select & "        CAST(ZM.板厚 AS VARCHAR) + 'X' +"
    mSL_Select = mSL_Select & "        CAST(ZM.幅   AS VARCHAR) + 'X' +"
    mSL_Select = mSL_Select & "        CAST(ZM.長さ AS VARCHAR) 材料"
    mSL_Select = mSL_Select & "   FROM 材料取引先マスタ ZT"
    mSL_Select = mSL_Select & "        LEFT OUTER JOIN 材料マスタ ZM"
    mSL_Select = mSL_Select & "             ON ZT.材料管理番号 = ZM.材料管理番号"
    mSL_Select = mSL_Select & "  WHERE ZT.背番号 = '" & RTrim(RV_Seban) & "'"
    mSL_Select = mSL_Select & "    AND ZT.取引先 = '" & RTrim(RV_Torcd) & "'"
'    If RVI_Hinmoku <> 9 Then
'       mSL_Select = mSL_Select & "    AND 品目 = " & RVI_Hinmoku
'    End If
    mSL_Select = mSL_Select & "  ORDER BY ZT.材質, ZT.受給品番, ZT.発注品番"
    
    Adodc1.RecordSource = mSL_Select    ' Select文セット
        
    Call Adodc1.Refresh
    
    H1lb_Tornm = RTrim(RV_TorRName)     ' 取引先名
    H1lb_Seban = RTrim(RV_Seban)        ' 背番号
    
    If Adodc1.Recordset.RecordCount = 0 Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(1) = "    該当するデータはありません  "
            .MB_Button = OK
            .MBOX
        End With
        
        Call PB_CAN_Click
        Exit Sub
    End If
    
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
    
'   ( 入力(Enter) )
    If PB_ENT.Visible Then
       If Key_Code = vbKeyReturn Then
           Key_Acc = True
           If IVENT = W_IVENT Then
               Dummy.Enabled = True
               Dummy.SetFocus
               DoEvents
               Call PB_ENT_Click
               Dummy.Enabled = False
           End If
           Call PB_CAN_Click
           Exit Function
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
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Me.Hide
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call RtnSet
End Sub
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TDBGridPro)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+      ADODC
'+----------------------------------------------------------+
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Adodc1.RecordSource <> "" Then Call TotalDisp
End Sub
'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub TotalDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    Me.Caption = MeCap & Space(5)
    Me.Caption = Me.Caption & "抽出件数 = " & Format(NowCnt, "##,##0")
    Me.Caption = Me.Caption & "/" & Format(MaxCnt, "##,##0") & " 件"
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SendKeys "{RIGHT}"
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If DBGrid1.ColContaining(X) < 0 Then Exit Sub
    DBGrid1.Row = DBGrid1.RowContaining(Y)
    DoEvents
    Call RtnSet
End Sub
'+-------------------------------------+
'+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
'+-------------------------------------+
Private Sub RtnSet()
    With Adodc1
        RV_Zaist = .Recordset("材質")
        RV_JHinbn = .Recordset("受給品番")
        RV_HHinbn = .Recordset("発注品番")
        Rv_KanriNo = .Recordset("材料管理番号")
        RV_Itatu = .Recordset("板厚")
        RV_Width = .Recordset("幅")
        RV_Long = .Recordset("長さ")
    End With
    
    Adodc1.Recordset.Close
    
    RV_Rtn = True
    Call PB_CAN_Click
End Sub

