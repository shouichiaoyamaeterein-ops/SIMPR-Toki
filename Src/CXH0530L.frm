VERSION 5.00
Begin VB.Form CXH0530L 
   Caption         =   "î[ïièëÅEéÛóÃèë"
   ClientHeight    =   11715
   ClientLeft      =   2550
   ClientTop       =   3510
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11715
   ScaleWidth      =   11775
   Begin VB.Image B1_Image1 
      Height          =   450
      Index           =   1
      Left            =   255
      Stretch         =   -1  'True
      Top             =   10845
      Width           =   450
   End
   Begin VB.Image B1_Image1 
      Height          =   450
      Index           =   0
      Left            =   255
      Stretch         =   -1  'True
      Top             =   5055
      Width           =   450
   End
   Begin VB.Label H1_Kbn 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "éÛ  óÃ  èë"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2085
      TabIndex        =   87
      Top             =   6030
      Width           =   1935
   End
   Begin VB.Label B1_LabelJS 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   765
      TabIndex        =   86
      Top             =   10965
      Width           =   2700
   End
   Begin VB.Label B1_NS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   6240
      TabIndex        =   85
      Top             =   10545
      Width           =   870
   End
   Begin VB.Label B1_HS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5400
      TabIndex        =   84
      Top             =   10545
      Width           =   750
   End
   Begin VB.Label B1_LabelGK 
      AutoSize        =   -1  'True
      Caption         =   "çáÅ@åv"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   83
      Top             =   10545
      Width           =   735
   End
   Begin VB.Line H1_K_LINE1 
      Index           =   0
      X1              =   6000
      X2              =   7155
      Y1              =   6675
      Y2              =   6675
   End
   Begin VB.Line H1_K_LINE1 
      Index           =   1
      X1              =   6000
      X2              =   7155
      Y1              =   7740
      Y2              =   7740
   End
   Begin VB.Line H1_K_LINE1 
      Index           =   2
      X1              =   6000
      X2              =   6000
      Y1              =   6675
      Y2              =   7740
   End
   Begin VB.Line H1_K_LINE1 
      Index           =   3
      X1              =   7155
      X2              =   7155
      Y1              =   6675
      Y2              =   7740
   End
   Begin VB.Label B1_LabelGyo 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   82
      Top             =   9855
      Width           =   255
   End
   Begin VB.Label B1_NS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   6240
      TabIndex        =   81
      Top             =   9825
      Width           =   870
   End
   Begin VB.Label B1_HS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   5400
      TabIndex        =   80
      Top             =   9825
      Width           =   750
   End
   Begin VB.Label B1_SS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   4320
      TabIndex        =   79
      Top             =   9825
      Width           =   855
   End
   Begin VB.Label B1_HM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   480
      TabIndex        =   78
      Top             =   10185
      Width           =   4815
   End
   Begin VB.Label B1_HB 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   480
      TabIndex        =   77
      Top             =   9825
      Width           =   3615
   End
   Begin VB.Label B1_LabelGyo 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   76
      Top             =   9135
      Width           =   255
   End
   Begin VB.Label B1_NS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   6240
      TabIndex        =   75
      Top             =   9105
      Width           =   870
   End
   Begin VB.Label B1_HS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5400
      TabIndex        =   74
      Top             =   9105
      Width           =   750
   End
   Begin VB.Label B1_SS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4320
      TabIndex        =   73
      Top             =   9105
      Width           =   855
   End
   Begin VB.Label B1_HM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   480
      TabIndex        =   72
      Top             =   9465
      Width           =   4815
   End
   Begin VB.Label B1_HB 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   480
      TabIndex        =   71
      Top             =   9105
      Width           =   3615
   End
   Begin VB.Label B1_LabelGyo 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   70
      Top             =   8430
      Width           =   255
   End
   Begin VB.Label H1_LabelNdH 
      Caption         =   "Å|"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2955
      TabIndex        =   69
      Top             =   7275
      Width           =   225
   End
   Begin VB.Line H1_LineNb 
      Index           =   1
      X1              =   3915
      X2              =   5670
      Y1              =   7515
      Y2              =   7515
   End
   Begin VB.Line H1_LineNd 
      Index           =   1
      X1              =   165
      X2              =   3825
      Y1              =   7515
      Y2              =   7515
   End
   Begin VB.Line H1_LineNouNum 
      Index           =   1
      X1              =   165
      X2              =   5670
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line H1_LineKbn2 
      Index           =   1
      X1              =   1800
      X2              =   4335
      Y1              =   6510
      Y2              =   6510
   End
   Begin VB.Line H1_LineKbn1 
      Index           =   1
      X1              =   1800
      X2              =   4335
      Y1              =   6465
      Y2              =   6465
   End
   Begin VB.Label H1_HKB 
      AutoSize        =   -1  'True
      Caption         =   "9999îN99åé99ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5700
      TabIndex        =   68
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label H1_LabeHKB 
      AutoSize        =   -1  'True
      Caption         =   "î≠çsì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   67
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label B1_LabelNS 
      Caption         =   "î[ì¸êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6510
      TabIndex        =   66
      Top             =   8025
      Width           =   750
   End
   Begin VB.Label B1_NS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   6240
      TabIndex        =   65
      Top             =   8385
      Width           =   870
   End
   Begin VB.Label B1_LabelHS 
      Caption         =   "î†  êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5550
      TabIndex        =   64
      Top             =   8025
      Width           =   750
   End
   Begin VB.Label B1_HS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5400
      TabIndex        =   63
      Top             =   8385
      Width           =   750
   End
   Begin VB.Label B1_LabelSS 
      Caption         =   "é˚óeêî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4590
      TabIndex        =   62
      Top             =   8025
      Width           =   735
   End
   Begin VB.Label B1_SS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4320
      TabIndex        =   61
      Top             =   8385
      Width           =   855
   End
   Begin VB.Label B1_HM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   60
      Top             =   8745
      Width           =   4815
   End
   Begin VB.Label B1_HB 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   59
      Top             =   8385
      Width           =   3615
   End
   Begin VB.Label B1_LabelHM 
      Caption         =   "ïi    ñº"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   525
      TabIndex        =   58
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label B1_LabelHB 
      Caption         =   "ïi    î‘"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   525
      TabIndex        =   57
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label H1_NouNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   225
      TabIndex        =   56
      Top             =   6840
      Width           =   4800
   End
   Begin VB.Label H1_LabelNouNum 
      Caption         =   "å‰íÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5130
      TabIndex        =   55
      Top             =   6825
      Width           =   555
   End
   Begin VB.Label H1_NdY 
      AutoSize        =   -1  'True
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   825
      TabIndex        =   54
      Top             =   7260
      Width           =   480
   End
   Begin VB.Label H1_NdM 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1710
      TabIndex        =   53
      Top             =   7260
      Width           =   240
   End
   Begin VB.Label H1_LabelNdT 
      Caption         =   "î[ì¸ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   52
      Top             =   7275
      Width           =   555
   End
   Begin VB.Label H1_NdD 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2355
      TabIndex        =   51
      Top             =   7260
      Width           =   240
   End
   Begin VB.Label H1_NdB 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3240
      TabIndex        =   50
      Top             =   7245
      Width           =   240
   End
   Begin VB.Label H1_LabelNdY 
      Caption         =   "îN"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1410
      TabIndex        =   49
      Top             =   7275
      Width           =   225
   End
   Begin VB.Label H1_LabelNdM 
      Caption         =   "åé"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2055
      TabIndex        =   48
      Top             =   7275
      Width           =   255
   End
   Begin VB.Label H1_LabelNdD 
      Caption         =   "ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2655
      TabIndex        =   47
      Top             =   7275
      Width           =   225
   End
   Begin VB.Label H1_LabelNdB 
      Caption         =   "ï÷"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   3570
      TabIndex        =   46
      Top             =   7275
      Width           =   255
   End
   Begin VB.Label H1_LabelNbT 
      Caption         =   "î[ì¸î‘çÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4020
      TabIndex        =   45
      Top             =   7275
      Width           =   810
   End
   Begin VB.Label H1_NbB 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4830
      TabIndex        =   44
      Top             =   7245
      Width           =   810
   End
   Begin VB.Label H1_NbB 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4830
      TabIndex        =   43
      Top             =   1455
      Width           =   810
   End
   Begin VB.Label H1_LabelNbT 
      Caption         =   "î[ì¸î‘çÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4020
      TabIndex        =   42
      Top             =   1485
      Width           =   870
   End
   Begin VB.Label H1_LabelNdB 
      Caption         =   "ï÷"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3570
      TabIndex        =   41
      Top             =   1485
      Width           =   255
   End
   Begin VB.Label H1_LabelNdD 
      Caption         =   "ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2655
      TabIndex        =   40
      Top             =   1485
      Width           =   225
   End
   Begin VB.Label H1_LabelNdM 
      Caption         =   "åé"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2055
      TabIndex        =   39
      Top             =   1485
      Width           =   255
   End
   Begin VB.Label H1_LabelNdY 
      Caption         =   "îN"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1410
      TabIndex        =   38
      Top             =   1485
      Width           =   225
   End
   Begin VB.Label H1_NdB 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   3240
      TabIndex        =   37
      Top             =   1470
      Width           =   240
   End
   Begin VB.Label H1_NdD 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2355
      TabIndex        =   36
      Top             =   1470
      Width           =   240
   End
   Begin VB.Label H1_LabelNdT 
      Caption         =   "î[ì¸ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   35
      Top             =   1485
      Width           =   555
   End
   Begin VB.Label H1_NdM 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   1710
      TabIndex        =   34
      Top             =   1470
      Width           =   240
   End
   Begin VB.Label H1_NdY 
      AutoSize        =   -1  'True
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   825
      TabIndex        =   33
      Top             =   1470
      Width           =   480
   End
   Begin VB.Label H1_LabelNouNum 
      Caption         =   "å‰íÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5130
      TabIndex        =   32
      Top             =   1035
      Width           =   555
   End
   Begin VB.Label H1_NouNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   225
      TabIndex        =   31
      Top             =   1050
      Width           =   4800
   End
   Begin VB.Label B1_LabelHB 
      Caption         =   "ïi    î‘"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   525
      TabIndex        =   30
      Top             =   2010
      Width           =   1095
   End
   Begin VB.Label B1_LabelHM 
      Caption         =   "ïi    ñº"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   525
      TabIndex        =   29
      Top             =   2250
      Width           =   1095
   End
   Begin VB.Label B1_HB 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   28
      Top             =   2595
      Width           =   3615
   End
   Begin VB.Label B1_HM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   27
      Top             =   2955
      Width           =   4815
   End
   Begin VB.Label B1_SS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   26
      Top             =   2595
      Width           =   855
   End
   Begin VB.Label B1_LabelSS 
      Caption         =   "é˚óeêî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4590
      TabIndex        =   25
      Top             =   2235
      Width           =   735
   End
   Begin VB.Label B1_HS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   24
      Top             =   2595
      Width           =   750
   End
   Begin VB.Label B1_LabelHS 
      Caption         =   "î†  êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5550
      TabIndex        =   23
      Top             =   2235
      Width           =   750
   End
   Begin VB.Label B1_NS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   22
      Top             =   2595
      Width           =   870
   End
   Begin VB.Label B1_LabelNS 
      Caption         =   "î[ì¸êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6510
      TabIndex        =   21
      Top             =   2235
      Width           =   750
   End
   Begin VB.Label H1_LabeHKB 
      AutoSize        =   -1  'True
      Caption         =   "î≠çsì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   20
      Top             =   450
      Width           =   615
   End
   Begin VB.Label H1_HKB 
      AutoSize        =   -1  'True
      Caption         =   "9999îN99åé99ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5700
      TabIndex        =   19
      Top             =   450
      Width           =   1575
   End
   Begin VB.Line H1_LineKbn1 
      Index           =   0
      X1              =   1800
      X2              =   4335
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Line H1_LineKbn2 
      Index           =   0
      X1              =   1800
      X2              =   4335
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line H1_LineNouNum 
      Index           =   0
      X1              =   165
      X2              =   5670
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line H1_LineNd 
      Index           =   0
      X1              =   165
      X2              =   3825
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Line H1_LineNb 
      Index           =   0
      X1              =   3915
      X2              =   5670
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Label H1_LabelNdH 
      Caption         =   "Å|"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2955
      TabIndex        =   18
      Top             =   1485
      Width           =   225
   End
   Begin VB.Label B1_LabelGyo 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label B1_HB 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   16
      Top             =   3315
      Width           =   3615
   End
   Begin VB.Label B1_HM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   15
      Top             =   3675
      Width           =   4815
   End
   Begin VB.Label B1_SS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   14
      Top             =   3315
      Width           =   855
   End
   Begin VB.Label B1_HS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   13
      Top             =   3315
      Width           =   750
   End
   Begin VB.Label B1_NS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   12
      Top             =   3315
      Width           =   870
   End
   Begin VB.Label B1_LabelGyo 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   3345
      Width           =   255
   End
   Begin VB.Label B1_HB 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Top             =   4035
      Width           =   3615
   End
   Begin VB.Label B1_HM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   4395
      Width           =   4815
   End
   Begin VB.Label B1_SS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   8
      Top             =   4035
      Width           =   855
   End
   Begin VB.Label B1_HS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   7
      Top             =   4035
      Width           =   750
   End
   Begin VB.Label B1_NS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   6
      Top             =   4035
      Width           =   870
   End
   Begin VB.Label B1_LabelGyo 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   4065
      Width           =   255
   End
   Begin VB.Line H1_K_LINE0 
      Index           =   3
      X1              =   7155
      X2              =   7155
      Y1              =   885
      Y2              =   1950
   End
   Begin VB.Line H1_K_LINE0 
      Index           =   2
      X1              =   6000
      X2              =   6000
      Y1              =   885
      Y2              =   1950
   End
   Begin VB.Line H1_K_LINE0 
      Index           =   1
      X1              =   6000
      X2              =   7155
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line H1_K_LINE0 
      Index           =   0
      X1              =   6000
      X2              =   7155
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Label B1_LabelGK 
      AutoSize        =   -1  'True
      Caption         =   "çáÅ@åv"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   4
      Top             =   4755
      Width           =   735
   End
   Begin VB.Label B1_HS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   3
      Top             =   4755
      Width           =   750
   End
   Begin VB.Label B1_NS 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   2
      Top             =   4755
      Width           =   870
   End
   Begin VB.Label B1_LabelJS 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   765
      TabIndex        =   1
      Top             =   5175
      Width           =   2700
   End
   Begin VB.Label H1_Kbn 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "î[  ïi  èë"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2085
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Line B1_LineC 
      BorderStyle     =   3  'ì_ê¸
      BorderWidth     =   2
      X1              =   240
      X2              =   7300
      Y1              =   5745
      Y2              =   5745
   End
End
Attribute VB_Name = "CXH0530L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  î≠íçñæç◊èëÅ@î≠çs
'**       ÉtÉHÅ[ÉÄID    :  CXH0530L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2008/08/20
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
'#î‰ärëﬁîóp#
    Dim Key_MeiNo       As Long
    Dim Key_Nouno       As String
    Dim Key_NPage       As Byte
    Dim Key_Siire       As String
    Dim Key_SKKbn       As Byte
    Dim GPMaisu         As Long         'ï≈ì‡î†êîçáåv
    Dim GTMaisu         As Long         'ñæç◊áÇî†êîçáåv
    Dim GPKensu         As Long         'ï≈ì‡åèêî
    Dim GTKensu         As Long         'ñæç◊áÇåèêî
    Dim TotalPage       As Long         'édì¸êÊëçï≈êî
    Dim SiirePage       As Long         'édì¸êÊåªï≈êî
    Dim MeiPage         As Long         'ñæç◊èëï≈êî
    
    Dim mPrtDate        As String
    Dim mPrtTime        As String
    Dim mJisyaNm        As String
    Dim mJisyaKj        As String
    
    Dim mImageFld       As String

'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyyîNmmåéddì˙")          'î≠çsì˙
    mPrtTime = Format(Now(), "hh:nn")                   'î≠çséûä‘
    
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrtCXH0530
        .iPrtTitle = "î[ïièëÅEéÛóÃèë"
        .iLineMax = 55
        .iLinePitch = 6
        .iPaperSize = p_A5 'p_A4
'       .iOrientation = po_PORTRAIT
        
        If Not .PrintIni Then
           With CisFun
                .MB_Lines = 5
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
    
'ÉçÉSImage
    Call GetClientOption("G", "Option", "Image", mImageFld)
    If RTrim(mImageFld) = "" Then
        Call GetClientOption("P", "Option", "Image", gIniExe)
    End If
    
    If Right(RTrim(mImageFld), 1) <> "\" Then mImageFld = mImageFld & "\"
    mImageFld = mImageFld & "logo.jpg"
    
    Call PrintProc
    
End Sub
'*---------------------------------------------------------------*
'*Å@ << àÛ  ç¸  èà  óù >>
'*---------------------------------------------------------------*
Private Function PrintProc()
    Dim Cnt_W1, Cnt_W2, Cnt_NB As Integer   ' àÛç¸ópÉJÉEÉìÉ^Å@2008.08.22 add

'*---------------------------------*
'*  é©é–ñºèÃÅEçHèÍñºÇãÅÇﬂÇÈ
'*---------------------------------*
' ÉVÉXÉeÉÄèÓïÒ(é©é–çHèÍ)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    Call IRNRead("", 1, 1)
' éÊà¯êÊÉ}ÉXÉ^ÇÊÇËñºèÃÇãÅÇﬂÇÈ
    gSL_Select = "SELECT * FROM éÊà¯êÊÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = 1"
    gSL_Select = gSL_Select & "   AND éÊà¯êÊCD = '" & RTrim(gCompany) & Left(IRN.oyian, 1) & "'"
    If TRMRead(gSL_Select, 1, 1) Then
        mJisyaNm = TRM.ê≥éÆñºèÃ '.ó™èÃ
        mJisyaKj = TRM.çHèÍñº
    End If
    
''*---------------------------------*
''*  èâä˙ê›íË
''*---------------------------------*
'    GPMaisu = 0: GPMaisu = 0: GPKensu = 0: GPKensu = 0
'    TotalPage = 0: SiirePage = 0: MeiPage = 0
'    Key_MeiNo = 0: Key_Nouno = "": Key_NPage = 0: Key_SKKbn = 0
'    Key_Siire = "": GTKensu = 0: GTMaisu = 0 ' 2004/05/10 Add
'    Call CisPrtCXH0530.ClrLstFld("B1")
'
'*-----------------------------------------*
'*  í†ï[èoóÕÉeÅ[ÉuÉãì«çûÅAàÛéöÉfÅ[É^çÏê¨
'*-----------------------------------------*
    B1_HS(0) = 0: B1_HS(10) = 0
    B1_NS(0) = 0: B1_NS(10) = 0
    Cnt_NB = 1
    
    Do Until Not CHT_RDSTS
    
        With CHT
            gSL_Select = "Select * From é¿ê—ä«óùÉeÅ[ÉuÉã " & vbCrLf
            gSL_Select = gSL_Select & " Where é¿ê—ä«óùNO = '" & .ä«óùNO & "' " & vbCrLf
            gSL_Select = gSL_Select & "   And î[ì¸î‘çÜ   = '" & .åüçıèÓïÒ & "' " & vbCrLf
'Call Test_LogOutput(gSL_Select, 1)
            
            Call ItemsClearJKT
            Call JKTRead(gSL_Select, 1, 1)
            With JKT
                gSL_Select = "Select ïiñº, ï\é¶ïiî‘ "
                gSL_Select = gSL_Select & " From  ïiî‘É}ÉXÉ^ "
                gSL_Select = gSL_Select & " Where ïiî‘   = '" & .ïiî‘ & "' "
                Call ItemsClearHNM
                Call HNMRead(gSL_Select, 1, 1)
    
                gSL_Select = "Select ê≥éÆñºèÃ, ó™èÃ "
                gSL_Select = gSL_Select & " From  éÊà¯êÊÉ}ÉXÉ^ "
                gSL_Select = gSL_Select & " Where éÊà¯êÊCD   = '" & .é¿ê—êÊ & "' "
                gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = 1 "
                Call ItemsClearTRM
                Call TRMRead(gSL_Select, 1, 1)
                
                For Cnt_W1 = 0 To 1
                    Cnt_W2 = (Cnt_W1 * 10) + Cnt_NB
                    H1_HKB(Cnt_W1) = mPrtDate
                    H1_NouNm(Cnt_W1) = TRM.ó™èÃ
                    H1_NdY(Cnt_W1) = Mid(.é¿ê—ì˙, 1, 4)
                    H1_NdM(Cnt_W1) = Mid(.é¿ê—ì˙, 5, 2)
                    H1_NdD(Cnt_W1) = Mid(.é¿ê—ì˙, 7, 2)
                    H1_NdB(Cnt_W1) = .é¿ê—ï÷
                    H1_NbB(Cnt_W1) = CHT.åüçıèÓïÒ
                    
                    B1_HB(Cnt_W2) = HNM.ï\é¶ïiî‘
                    B1_HM(Cnt_W2) = HNM.ïiñº
                
                    B1_SS(Cnt_W2) = Format(.é˚óeêî, "#,###")
                    B1_HS(Cnt_W2) = Format(.é¿ê—ñáêî, "#,###")
                    B1_NS(Cnt_W2) = Format(.é¿ê—êî, "#,###")
                    B1_HS(Cnt_W1 * 10) = B1_HS(Cnt_W1 * 10) + .é¿ê—ñáêî
                    B1_NS(Cnt_W1 * 10) = B1_NS(Cnt_W1 * 10) + .é¿ê—êî
                    B1_LabelJS(Cnt_W1) = mJisyaNm
                    B1_Image1(Cnt_W1).Stretch = True
                    'B1_Image1(Cnt_W1).Picture = LoadPicture("../Exe/BMP/logo.jpg")
                    B1_Image1(Cnt_W1).Picture = LoadPicture(mImageFld)
                Next Cnt_W1
            End With
            Cnt_NB = Cnt_NB + 1
            Call CHTReadNext
            
            If Cnt_NB > 3 Or .åüçıèÓïÒ <> H1_NbB(0) Or CHT_RDSTS = False Then
                If Cnt_NB <= 3 Then
                    Do Until Cnt_NB > 3
                        For Cnt_W1 = 0 To 1                 ' àÛéöÉfÅ[É^Ç™Ç»Ç¢ïîï™ÇãÛîíÇ…Ç∑ÇÈ
                            Cnt_W2 = (Cnt_W1 * 10) + Cnt_NB
                            B1_HB(Cnt_W2) = ""
                            B1_HM(Cnt_W2) = ""
                            B1_SS(Cnt_W2) = ""
                            B1_HS(Cnt_W2) = ""
                            B1_NS(Cnt_W2) = ""
                        Next Cnt_W1
                        Cnt_NB = Cnt_NB + 1
                    Loop
                End If
                
                B1_HS(0) = Format(B1_HS(0), "#,###")        ' çáåvÉtÉHÅ[É}ÉbÉgê›íË
                B1_NS(0) = Format(B1_NS(0), "#,###")
                B1_HS(10) = Format(B1_HS(10), "#,###")
                B1_NS(10) = Format(B1_NS(10), "#,###")
                
                '*-- ÉGÉâÅ[ÉPÅ[ÉX --*
'                If Not JKT_RDSTS Then
'                    H1_NouNm(0) = ""
'                    H1_NouNm(1) = ""
'                    B1_HB(2) = "ñ¢ìoò^ÇÃî[ì¸î‘çÜÅ@Ç‹ÇΩÇÕ"
'                    B1_HM(2) = "ï ÇÃí[ññÇ≈ìoò^Ç≥ÇÍÇΩî[ì¸î‘çÜÇ≈Ç∑"
'                End If
                
                With CisPrtCXH0530
                    .cDataCnt = .cDataCnt + 1
                    .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXH0530_PrtCnt, "#,###")
                    .cLine = 0
                    .cPage = .cPage + 1
                   '.ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
                    .PrintNewPage       'â¸ï≈èàóù
                    .PrintFld ("H1")    'HEADïîàÛç¸
                    .PrintFld ("B1")
                End With
                
                B1_HS(0) = 0: B1_HS(10) = 0
                B1_NS(0) = 0: B1_NS(10) = 0
                Cnt_NB = 1
            End If
        End With
    Loop
    
    Exit Function

CoPrintH0530_Err:
    CisFun.ErrorBox
    End
End Function

