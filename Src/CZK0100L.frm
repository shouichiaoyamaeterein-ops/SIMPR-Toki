VERSION 5.00
Begin VB.Form CZK0100L 
   Caption         =   "íIâµÉäÉXÉg"
   ClientHeight    =   4485
   ClientLeft      =   1860
   ClientTop       =   3165
   ClientWidth     =   15705
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   15705
   Begin VB.Line B2_Line 
      BorderStyle     =   3  'ì_ê¸
      X1              =   0
      X2              =   16000
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Label B4_Kingaku 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "###,###,##0"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12615
      TabIndex        =   41
      Top             =   1905
      Width           =   1320
   End
   Begin VB.Label B4_Ko 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "å¬"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10470
      TabIndex        =   40
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label B4_Tanasu 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "#,###,##0"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9375
      TabIndex        =   39
      Top             =   1905
      Width           =   1080
   End
   Begin VB.Label B4_SumJB 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "[ é©ÉKÉX  åv ]"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6360
      TabIndex        =   38
      Top             =   1905
      Width           =   1695
   End
   Begin VB.Label H1_JiGass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "XXXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   37
      Top             =   720
      Width           =   735
   End
   Begin VB.Line B4_Line 
      BorderStyle     =   3  'ì_ê¸
      X1              =   30
      X2              =   16030
      Y1              =   2145
      Y2              =   2145
   End
   Begin VB.Label B3_Kingaku 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "###,###,##0"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12615
      TabIndex        =   36
      Top             =   2430
      Width           =   1320
   End
   Begin VB.Label B3_Ko 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "å¬"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10470
      TabIndex        =   35
      Top             =   2415
      Width           =   240
   End
   Begin VB.Label B3_Tanasu 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "#,###,##0"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9375
      TabIndex        =   34
      Top             =   2430
      Width           =   1080
   End
   Begin VB.Label B3_SumS 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "[ ëç  çá  åv ]"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6360
      TabIndex        =   33
      Top             =   2430
      Width           =   1695
   End
   Begin VB.Label B2_Kingaku 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "###,###,##0"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12615
      TabIndex        =   32
      Top             =   2175
      Width           =   1320
   End
   Begin VB.Label B2_Ko 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "å¬"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10470
      TabIndex        =   31
      Top             =   2175
      Width           =   240
   End
   Begin VB.Label B2_Tanasu 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "#,###,##0"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9375
      TabIndex        =   30
      Top             =   2175
      Width           =   1080
   End
   Begin VB.Label B2_SumB 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "[ ïîïiãÊï™åv ]"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6360
      TabIndex        =   29
      Top             =   2175
      Width           =   1695
   End
   Begin VB.Label B1_Tanasu 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "#,###,##0"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9375
      TabIndex        =   28
      Top             =   1665
      Width           =   1080
   End
   Begin VB.Label B1_Ko 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "å¬"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10470
      TabIndex        =   27
      Top             =   1665
      Width           =   240
   End
   Begin VB.Label B1_Tanka 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "##,###,##0.00"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10935
      TabIndex        =   26
      Top             =   1665
      Width           =   1560
   End
   Begin VB.Label B1_Kingaku 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "###,###,##0"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12615
      TabIndex        =   25
      Top             =   1665
      Width           =   1320
   End
   Begin VB.Label B1_Shiyo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6360
      TabIndex        =   24
      Top             =   1665
      Width           =   3000
   End
   Begin VB.Label B1_Saisyu 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "YYYY/MM/DD"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   14295
      TabIndex        =   23
      Top             =   1665
      Width           =   1200
   End
   Begin VB.Line B1_LINE 
      BorderStyle     =   3  'ì_ê¸
      X1              =   0
      X2              =   16000
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ç≈èIèoå…ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   14280
      TabIndex        =   22
      Top             =   1395
      Width           =   1215
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   6780
      TabIndex        =   21
      Top             =   5235
      Width           =   120
   End
   Begin VB.Label H1_TanaDay 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   795
      Width           =   3015
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "édÅ@Å@ól"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6960
      TabIndex        =   19
      Top             =   1395
      Width           =   855
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ïiÅ@Å@î‘"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   18
      Top             =   1395
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "éËîzêÊÅ@ÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   17
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label H1_ToriNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ïîïiãÊï™ÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label H1_HKubun 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   14
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label H1_ZPlace 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ç›å…èÍèäÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ã‡äz"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   13440
      TabIndex        =   11
      Top             =   1395
      Width           =   495
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "íPâø"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   10
      Top             =   1395
      Width           =   495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "íIâµêî"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   9
      Top             =   1395
      Width           =   735
   End
   Begin VB.Label B1_Hinmei 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3360
      TabIndex        =   8
      Top             =   1665
      Width           =   3000
   End
   Begin VB.Label B1_Hinban 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   1665
      Width           =   3000
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ïiÅ@Å@ñº"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   6
      Top             =   1395
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "(CZK0100L)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "Å¶Å@Å@íIÅ@Å@âµÅ@Å@ÉäÅ@Å@ÉXÅ@Å@ÉgÅ@Å@Å¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label H1_Page 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "9,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   15150
      TabIndex        =   3
      Top             =   345
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "Page:"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14760
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "î≠çsì˙:"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12600
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Label H1_Date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "9999/99/99 "
      Height          =   255
      Left            =   13260
      TabIndex        =   0
      Top             =   345
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15985
      Y1              =   1635
      Y2              =   1635
   End
End
Attribute VB_Name = "CZK0100L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  íIâµÉäÉXÉgÅ@î≠çs
'**       ÉtÉHÅ[ÉÄID    :  CZK0100L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2010/01/07
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim mPrtDate        As String
    
    Dim wPlace          As String
    Dim wKubun          As String
    Dim wJiGass         As Byte
    Dim wTori           As String
    Dim wTanaDay        As String
    Dim wGyoCnt         As Long

    Dim wTanaB          As Long
    Dim wTanaS          As Long
    Dim wTanaJG         As Long
    Dim wKinB           As Long
    Dim wKinS           As Long
    Dim wKinJG          As Long
    
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg   *
'*------------------------------------*
Private Sub form_load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             'î≠çsì˙
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrtCZK0100
        .iPrtTitle = "íIâµÉäÉXÉg"
        .iLineMax = 40
        .iLineInterval = 0
        .iLinePitch = 6
        .iPaperSize = p_A4
        .iOrientation = po_Landscape
        
        If Not .PrintIni Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "     Å@ÉvÉäÉìÉ^èâä˙ê›íËÉGÉâÅ[        "
                .MB_Title = "àÛç¸¥◊∞"
                .MB_Button = Error
                .MBOX
           End With
           Exit Sub
        .cLine = 0
        .cPage = 0
        End If
    End With

    Call PrintProc

End Sub
'*---------------------------------------------------------------*
'*Å@ << àÛ  ç¸  èà  óù >>
'*---------------------------------------------------------------*
Private Function PrintProc()
'*---------------------------------*
'*  èâä˙ê›íË
'*---------------------------------*
    wPlace = ""
    wKubun = ""
    wTori = ""
    wTanaDay = ""
    wGyoCnt = 0
    wTanaS = 0
    wTanaB = 0
    wTanaJG = 0
    wKinS = 0
    wKinB = 0
    wKinJG = 0
    wJiGass = 255
'*---------------------------------*
'*       ç›å…É}ÉXÉ^ì«çûÇ›          *
'*---------------------------------*
    Do Until Not ZKM_RDSTS
        If wPlace <> ZKM.ç›å…èÍèä Or wKubun <> ZKM.ïiñ⁄ãÊï™ Or wJiGass <> ZKM.é©ÉKÉX Or wTori <> ZKM.éÊà¯êÊ Or wTanaDay <> ZKM.é¿íIì˙ Then
            wPlace = ZKM.ç›å…èÍèä
            wKubun = ZKM.ïiñ⁄ãÊï™
            wJiGass = ZKM.é©ÉKÉX
            wTori = ZKM.éÊà¯êÊ
            wTanaDay = ZKM.é¿íIì˙
            wGyoCnt = 0
            CisPrtCZK0100.cLine = CisPrtCZK0100.iLineMax
        End If

        Call HeadWrite(0)

        With CisPrtCZK0100
            .ClrLstFld ("B1")
        End With
        '*---------------------------------*
        '*  BodyçsÅ@ï“èW
        '*---------------------------------*
        With ZKM
            B1_Hinban = .ïiî‘
            B1_Hinmei = .ïiñº
            B1_Shiyo = .édól
            B1_Tanasu = Format(.ç›å…êî, "#,##0")
            B1_Ko = "å¬"
            B1_Tanka = Format(.íPâø, "#,##0.00")
            B1_Kingaku = Format(.ã‡äz, "#,##0")
            If .ç≈èIèoå…ì˙ = "" Then
                B1_Saisyu = ""
            Else
                B1_Saisyu = Mid(.ç≈èIèoå…ì˙, 1, 4) & "/" & Mid(.ç≈èIèoå…ì˙, 5, 2) & "/" & Mid(.ç≈èIèoå…ì˙, 7, 2)
            End If
            wGyoCnt = wGyoCnt + 1
            wTanaB = wTanaB + .ç›å…êî
            wKinB = wKinB + .ã‡äz
            wTanaJG = wTanaJG + .ç›å…êî
            wKinJG = wKinJG + .ã‡äz
            
            If Fix(wGyoCnt) Mod 2 = 0 Then
                If Not wGyoCnt = 40 Then
                    B1_LINE.Visible = True
                Else
                    B1_LINE.Visible = False
                    wGyoCnt = 0
                End If
            Else
                B1_LINE.Visible = False
            End If
            
            Call ZKMReadNext
            
            If (wPlace = ZKM.ç›å…èÍèä And wKubun = ZKM.ïiñ⁄ãÊï™ And wJiGass = ZKM.é©ÉKÉX And (wTori <> ZKM.éÊà¯êÊ Or wTanaDay <> ZKM.é¿íIì˙)) And ZKM_RDSTS Then
                B1_LINE.Visible = False
            End If
        End With
        '*---------------------------------*
        '*  BodyïîÅ@àÛç¸
        '*---------------------------------*
        With CisPrtCZK0100
            .PrintFld ("B1*")
            .cLine = .cLine + 1
    '       ÉJÉEÉìÉ^Å@ÇtÇo
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With
        
        If wPlace <> ZKM.ç›å…èÍèä Or wKubun <> ZKM.ïiñ⁄ãÊï™ Or Not ZKM_RDSTS Then
            Call HeadWrite(0)
        
            With CisPrtCZK0100
                .ClrLstFld ("B2")
                B2_SumB = "[ ïîïiãÊï™åv ]"
                B2_Tanasu = Format(wTanaB, "#,##0")
                B2_Kingaku = Format(wKinB, "#,##0")
                B2_Ko = "å¬"
                wGyoCnt = wGyoCnt + 1
                wTanaS = wTanaS + wTanaB
                wKinS = wKinS + wKinB
                wTanaB = 0
                wKinB = 0
                If Not ZKM_RDSTS Then
                    If Fix(wGyoCnt) Mod 2 = 0 Then
                        If Not wGyoCnt = 40 Then
                            B2_Line.Visible = True
                        Else
                            B2_Line.Visible = False
                            wGyoCnt = 0
                        End If
                    Else
                        B2_Line.Visible = False
                    End If
                Else
                    B2_Line.Visible = False
                End If
                '*---------------------------------*
                '*  Body2ïîÅ@àÛç¸
                '*---------------------------------*
                If Not .cLine = 0 Then .cLine = .cLine - 1
                .PrintFld ("B2*")
                .cLine = .cLine + 1
            End With
        End If
        
        If wPlace <> ZKM.ç›å…èÍèä Or wKubun <> ZKM.ïiñ⁄ãÊï™ Or wJiGass <> ZKM.é©ÉKÉX Or Not ZKM_RDSTS Then
            Call HeadWrite(0)
            
            With CisPrtCZK0100
                .ClrLstFld ("B4")
                If wJiGass = 0 Then
                    B4_SumJB = "[  é©ìÆé‘åv  ]"
                Else
                    B4_SumJB = "[   ÉKÉXåv   ]"
                End If
                B4_Tanasu = Format(wTanaJG, "#,##0")
                B4_Kingaku = Format(wKinJG, "#,##0")
                B4_Ko = "å¬"
                wGyoCnt = wGyoCnt + 1
                wTanaJG = 0
                wKinJG = 0
                If Not ZKM_RDSTS Then
                    If Fix(wGyoCnt) Mod 2 = 0 Then
                        If Not wGyoCnt = 40 Then
                            B4_Line.Visible = True
                        Else
                            B4_Line.Visible = False
                            wGyoCnt = 0
                        End If
                    Else
                        B4_Line.Visible = False
                    End If
                Else
                    B4_Line.Visible = False
                End If
                '*---------------------------------*
                '*  Body4ïîÅ@àÛç¸
                '*---------------------------------*
                If Not .cLine = 0 Then .cLine = .cLine - 1
                .PrintFld ("B4*")
                .cLine = .cLine + 1
        
            End With
                    
        End If
        
    Loop
    Call HeadWrite(0)
    With CisPrtCZK0100
        .ClrLstFld ("B3")
        B3_SumS = "[ ëç  çá  åv ]"
        B3_Tanasu = Format(wTanaS, "#,##0")
        B3_Kingaku = Format(wKinS, "#,##0")
        B3_Ko = "å¬"
        wGyoCnt = wGyoCnt + 1
        '*---------------------------------*
        '*  Body3ïîÅ@àÛç¸
        '*---------------------------------*
        If Not .cLine = 0 Then .cLine = .cLine - 1
        .PrintFld ("B3*")
        .cLine = .cLine + 1
    End With
    Call ZKMClose(1)

    ' àÛç¸èIóπ
    Call CisPrtCZK0100.PrintEndDoc

End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False

    With CisPrtCZK0100
    '   ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function

        wGyoCnt = 0
        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé

    '   î≠çsì˙ÅEï≈
        H1_Date = mPrtDate
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")
        
        H1_ZPlace = wPlace
        H1_HKubun = wKubun
        If wJiGass = 0 Then
            H1_JiGass = "é©ìÆé‘"
        Else
            H1_JiGass = "ÉKÉX"
        End If
        
        H1_ToriNm = wTori
        If Len(wTanaDay) = 8 Then
            H1_TanaDay = "ÅyÅ@íIâµì˙  Å@ÅFÅ@" & Mid(wTanaDay, 1, 4) & "/" & Mid(wTanaDay, 5, 2) & "/" & Mid(wTanaDay, 7, 2) & "Å@Åz"
        Else
            H1_TanaDay = "ÅyÅ@íIâµîNåéÅ@ÅFÅ@" & Mid(wTanaDay, 1, 4) & "/" & Mid(wTanaDay, 5, 2) & "Å@Åz"
        End If

        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With

    HeadWrite = True
End Function
