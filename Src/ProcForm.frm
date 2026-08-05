VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Begin VB.Form ProcForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   1620
   ClientLeft      =   75
   ClientTop       =   1500
   ClientWidth     =   1770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   Begin Cis3D_v60.CIS3D Proc_lbl 
      Height          =   300
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      Caption         =   "  複写 (Alt+C)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   12632256
      cBoderColor2    =   12632256
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cAlingnment     =   6
      cBoderStyle     =   2
      cPositionY      =   15
   End
   Begin VB.PictureBox Picture1 
      Height          =   75
      Left            =   -120
      ScaleHeight     =   15
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   1230
      Width           =   2595
   End
   Begin Cis3D_v60.CIS3D Proc_lbl 
      Height          =   300
      Index           =   1
      Left            =   0
      Top             =   300
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      Caption         =   "  挿入 (Alt+E)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   12632256
      cBoderColor2    =   12632256
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cAlingnment     =   6
      cBoderStyle     =   2
      cPositionY      =   15
   End
   Begin Cis3D_v60.CIS3D Proc_lbl 
      Height          =   300
      Index           =   2
      Left            =   0
      Top             =   600
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      Caption         =   "  移動 (Alt+F)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   12632256
      cBoderColor2    =   12632256
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cAlingnment     =   6
      cBoderStyle     =   2
      cPositionY      =   15
   End
   Begin Cis3D_v60.CIS3D Proc_lbl 
      Height          =   300
      Index           =   3
      Left            =   0
      Top             =   900
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      Caption         =   "  削除 (Alt+D)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   12632256
      cBoderColor2    =   12632256
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cAlingnment     =   6
      cBoderStyle     =   2
      cPositionY      =   15
   End
   Begin Cis3D_v60.CIS3D Proc_lbl 
      Height          =   300
      Index           =   4
      Left            =   0
      Top             =   1320
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      Caption         =   "  戻   る (Esc)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   12632256
      cBoderColor2    =   12632256
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cAlingnment     =   6
      cBoderStyle     =   2
      cPositionY      =   15
   End
End
Attribute VB_Name = "ProcForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  処理選択画面
'**       フォームID    :  ProcFrom
'**       処理概要      :
'**
'**       作  成  日    :  2004/03/22  By CIS
'**       変  更  日    :
'**
'**    gProcF_Left = 0
'**    gProcF_Top = 0
'**    gProcF_Enabled(1) = 1   '使用不可にするラベルを『１』に設定
'**
'**    ProcForm1.Show vbModal
'**
'**    Select Case gProcF_Rtn
'**        Case 0
'**        Case 1
'**    End Select
'**********************************************************'
    Dim Gyo_Sel     As Integer
    Dim wInt1       As Integer
    Dim wInt2       As Integer

    Const MaxIndex  As Integer = 4          ' 選択ラベルの最大Index№
    Const clr_Gray  As Long = &H8000000F    ' 非選択ラベルのBackColor
    Const clr_Bule  As Long = &HFFFFC0      '   選択ラベルのBackColor

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    Me.Move gProcF_Left, gProcF_Top, 1890

    Gyo_Sel = -1
End Sub
'+-------------------------------------------------------+
'+      フォーム    《 ＫｅｙＵｐイベント 》               +
'+-------------------------------------------------------+
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyReturn And KeyCode <> vbKeyEscape Then Exit Sub
    
'   *------------------------*
'   * Return Key での選択終了 *
'   *------------------------*
    If KeyCode = vbKeyReturn Then
        If Gyo_Sel <> -1 Then Call Rtn_END
        Exit Sub
    End If
    
'   *------------------------*
'   * Escape Key での取消終了 *
'   *------------------------*
    If KeyCode = vbKeyEscape Then
        Call SelColor(MaxIndex)
        Call Rtn_END
        Exit Sub
    End If
    
'   *-----------------------------*
'   * Up & Down Key での選択行変更 *
'   *-----------------------------*
    For wInt1 = 1 To MaxIndex + 1
        If KeyCode = vbKeyUp Then
            wInt2 = Gyo_Sel - wInt1
        Else
            wInt2 = Gyo_Sel + wInt1
        End If
    
        If wInt2 >= 0 And wInt2 <= 4 Then
            Call ColorReturn
            Call SelColor(wInt2)
            Exit For
        End If
    Next wInt1
End Sub
'+------------------------------------------------------+
'+    Proc_lbl_MouseMove  * MouseMove での選択行変更 *   +
'+------------------------------------------------------+
Private Sub Proc_lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = Gyo_Sel Then Exit Sub
    Call ColorReturn
    Call SelColor(Index)
End Sub
'+-----------------------------------------------------+
'+    Proc_lbl_Click    * Mouse Click での選択終了 *    +
'+-----------------------------------------------------+
Private Sub Proc_lbl_Click(Index As Integer)
    If Index <> Gyo_Sel Then Call ColorReturn
    Call SelColor(Index)
    Call Rtn_END
End Sub
'+-----------------------------+
'+    選択が外れた時の色設定     +
'+-----------------------------+
Private Sub ColorReturn()
    If Gyo_Sel <> -1 Then
        If Proc_lbl(Gyo_Sel).BackColor <> clr_Gray Then
            Proc_lbl(Gyo_Sel).BackColor = clr_Gray
            Proc_lbl(Gyo_Sel).cFont3D = 2
        End If
        Gyo_Sel = -1
    End If
End Sub
'+---------------------+
'+    選択設定 & 終了   +
'+---------------------+
Private Sub Rtn_END()
    gProcF_Rtn = Gyo_Sel
    Unload Me
End Sub
'+---------------------------+
'+    選択行のBackColor設定   +
'+---------------------------+
Private Sub SelColor(Index As Integer)
    Gyo_Sel = Index
    If Proc_lbl(Gyo_Sel).BackColor <> clr_Bule Then
        Proc_lbl(Gyo_Sel).BackColor = clr_Bule
        Proc_lbl(Gyo_Sel).cFont3D = 0
    End If
End Sub
