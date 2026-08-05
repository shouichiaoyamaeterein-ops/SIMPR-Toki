VERSION 5.00
Begin VB.Form COS0310L 
   Caption         =   "îÉä|ñæç◊ÉäÉXÉgÅiï°êîÅj"
   ClientHeight    =   3330
   ClientLeft      =   675
   ClientTop       =   2190
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   15420
   Begin VB.Label B1_Juryo 
      AutoSize        =   -1  'True
      Caption         =   "990.000"
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
      Left            =   13350
      TabIndex        =   56
      Top             =   2355
      Width           =   840
   End
   Begin VB.Label KTitle6 
      AutoSize        =   -1  'True
      Caption         =   "èdó "
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
      Left            =   13680
      TabIndex        =   55
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label B1_DD 
      AutoSize        =   -1  'True
      Caption         =   "DD"
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
      Index           =   8
      Left            =   19980
      TabIndex        =   54
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label B1_SU 
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
      Height          =   225
      Index           =   8
      Left            =   21240
      TabIndex        =   53
      Top             =   2355
      Width           =   840
   End
   Begin VB.Label B1_NO 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Index           =   8
      Left            =   20400
      TabIndex        =   52
      Top             =   2355
      Width           =   600
   End
   Begin VB.Label B1_DD 
      AutoSize        =   -1  'True
      Caption         =   "DD"
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
      Index           =   7
      Left            =   17400
      TabIndex        =   51
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label B1_SU 
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
      Height          =   225
      Index           =   7
      Left            =   18660
      TabIndex        =   50
      Top             =   2355
      Width           =   840
   End
   Begin VB.Label B1_NO 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Index           =   7
      Left            =   17820
      TabIndex        =   49
      Top             =   2355
      Width           =   600
   End
   Begin VB.Label B1_DD 
      AutoSize        =   -1  'True
      Caption         =   "DD"
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
      Index           =   6
      Left            =   14820
      TabIndex        =   48
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label B1_SU 
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
      Height          =   225
      Index           =   6
      Left            =   16080
      TabIndex        =   47
      Top             =   2355
      Width           =   840
   End
   Begin VB.Label B1_NO 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Index           =   6
      Left            =   15240
      TabIndex        =   46
      Top             =   2355
      Width           =   600
   End
   Begin VB.Label B1_MIDASI 
      AutoSize        =   -1  'True
      Caption         =   "Ås çá åv Åt"
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
      Left            =   1515
      TabIndex        =   45
      Top             =   2355
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "éÛì¸åé"
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
      Left            =   8910
      TabIndex        =   44
      Top             =   1905
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ã‡Å@äz"
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
      Left            =   7965
      TabIndex        =   43
      Top             =   1905
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "éÛì¸èÍèä"
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
      Left            =   2460
      TabIndex        =   42
      Top             =   1905
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ïiÅ@Å@î‘"
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
      Left            =   735
      TabIndex        =   41
      Top             =   1905
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "íPÅ@âø"
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
      Left            =   6405
      TabIndex        =   40
      Top             =   1905
      Width           =   675
   End
   Begin VB.Label B1_YM 
      AutoSize        =   -1  'True
      Caption         =   "MM"
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
      Left            =   9120
      TabIndex        =   39
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label B1_KINGK 
      AutoSize        =   -1  'True
      Caption         =   "999,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7320
      TabIndex        =   38
      Top             =   2355
      Width           =   1485
   End
   Begin VB.Label B1_Suryo 
      AutoSize        =   -1  'True
      Caption         =   "999,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3840
      TabIndex        =   37
      Top             =   2355
      Width           =   1485
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "êîÅ@ó "
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
      Left            =   4485
      TabIndex        =   36
      Top             =   1905
      Width           =   675
   End
   Begin VB.Label B1_TANKA 
      AutoSize        =   -1  'True
      Caption         =   "999,999,999.99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5385
      TabIndex        =   35
      Top             =   2355
      Width           =   1890
   End
   Begin VB.Label B1_Uke 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
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
      Left            =   2670
      TabIndex        =   34
      Top             =   2355
      Width           =   480
   End
   Begin VB.Label B1_HINBN 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX-XXXXX-XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   345
      TabIndex        =   33
      Top             =   2355
      Width           =   1890
   End
   Begin VB.Label H1_SIKINM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXXXXXXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1245
      TabIndex        =   32
      Top             =   1410
      Width           =   2895
   End
   Begin VB.Label B1_DD 
      AutoSize        =   -1  'True
      Caption         =   "DD"
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
      Left            =   1920
      TabIndex        =   31
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label B1_SU 
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
      Height          =   225
      Index           =   1
      Left            =   3180
      TabIndex        =   30
      Top             =   2355
      Width           =   840
   End
   Begin VB.Label B1_NO 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Left            =   2340
      TabIndex        =   29
      Top             =   2355
      Width           =   600
   End
   Begin VB.Label B1_NO 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Index           =   2
      Left            =   4920
      TabIndex        =   28
      Top             =   2355
      Width           =   600
   End
   Begin VB.Label B1_SU 
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
      Height          =   225
      Index           =   2
      Left            =   5760
      TabIndex        =   27
      Top             =   2355
      Width           =   840
   End
   Begin VB.Label B1_DD 
      AutoSize        =   -1  'True
      Caption         =   "DD"
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
      Index           =   2
      Left            =   4500
      TabIndex        =   26
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label B1_NO 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Index           =   3
      Left            =   7500
      TabIndex        =   25
      Top             =   2355
      Width           =   600
   End
   Begin VB.Label B1_SU 
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
      Height          =   225
      Index           =   3
      Left            =   8340
      TabIndex        =   24
      Top             =   2355
      Width           =   840
   End
   Begin VB.Label B1_DD 
      AutoSize        =   -1  'True
      Caption         =   "DD"
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
      Index           =   3
      Left            =   7080
      TabIndex        =   23
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label B1_NO 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Index           =   4
      Left            =   10080
      TabIndex        =   22
      Top             =   2355
      Width           =   600
   End
   Begin VB.Label B1_SU 
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
      Height          =   225
      Index           =   4
      Left            =   10920
      TabIndex        =   21
      Top             =   2355
      Width           =   840
   End
   Begin VB.Label B1_DD 
      AutoSize        =   -1  'True
      Caption         =   "DD"
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
      Index           =   4
      Left            =   9660
      TabIndex        =   20
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label B1_NO 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Index           =   5
      Left            =   12660
      TabIndex        =   19
      Top             =   2355
      Width           =   600
   End
   Begin VB.Label B1_SU 
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
      Height          =   225
      Index           =   5
      Left            =   13500
      TabIndex        =   18
      Top             =   2355
      Width           =   840
   End
   Begin VB.Label B1_DD 
      AutoSize        =   -1  'True
      Caption         =   "DD"
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
      Index           =   5
      Left            =   12240
      TabIndex        =   17
      Top             =   2355
      Width           =   240
   End
   Begin VB.Line H1_LINE 
      X1              =   315
      X2              =   22080
      Y1              =   2235
      Y2              =   2235
   End
   Begin VB.Label H1_SIKI 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   735
      TabIndex        =   16
      Top             =   1410
      Width           =   375
   End
   Begin VB.Label KTitle1 
      AutoSize        =   -1  'True
      Caption         =   "í˘ê≥óùóR"
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
      Left            =   11325
      TabIndex        =   15
      Top             =   1905
      Width           =   900
   End
   Begin VB.Label B1_Teisei 
      AutoSize        =   -1  'True
      Caption         =   "XX"
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
      Left            =   11685
      TabIndex        =   14
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label H1_TORCD 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   735
      TabIndex        =   13
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label H1_TORNM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXXXXXXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label KTitle5 
      AutoSize        =   -1  'True
      Caption         =   "ÉJÅ[Éh"
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
      Left            =   12390
      TabIndex        =   11
      Top             =   1905
      Width           =   675
   End
   Begin VB.Label KTitle4 
      AutoSize        =   -1  'True
      Caption         =   "ä«óùî‘çÜ"
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
      Left            =   9990
      TabIndex        =   10
      Top             =   1905
      Width           =   900
   End
   Begin VB.Label B1_Card 
      AutoSize        =   -1  'True
      Caption         =   "XX"
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
      Left            =   12645
      TabIndex        =   9
      Top             =   2355
      Width           =   240
   End
   Begin VB.Label B1_Kanri 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXXXX"
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
      Left            =   9810
      TabIndex        =   8
      Top             =   2355
      Width           =   1320
   End
   Begin VB.Label H1_MM 
      AutoSize        =   -1  'True
      Caption         =   "MM"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   720
      Width           =   300
   End
   Begin VB.Label H1_YY 
      AutoSize        =   -1  'True
      Caption         =   "YYYY"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "îN"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   1395
      TabIndex        =   5
      Top             =   720
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åéìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   2130
      TabIndex        =   4
      Top             =   705
      Width           =   570
   End
   Begin VB.Label H1_PAGE 
      AutoSize        =   -1  'True
      Caption         =   "9,999"
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
      Left            =   18090
      TabIndex        =   3
      Top             =   600
      Width           =   600
   End
   Begin VB.Label H1_YMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD HH:MM"
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
      Left            =   16170
      TabIndex        =   2
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ï≈"
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
      Index           =   1
      Left            =   18720
      TabIndex        =   1
      Top             =   600
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Å¶ îÉ ä| ã‡ ñæ ç◊ Éä ÉX Ég Å¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   8715
      TabIndex        =   0
      Top             =   600
      Width           =   5220
   End
End
Attribute VB_Name = "COS0310L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  îÉä|ã‡ñæç◊ÉäÉXÉgÅyï°êîÅz
'**       ÉtÉHÅ[ÉÄID    :  COS0310L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2004/11/05
'**       ïœ  çX  ì˙    :  2005/03/31 åvéZï˚ñ@ÇéÊà¯êÊÉ}ÉXÉ^ÇÃê›íËÇ…è]Ç§
'**       ïœ  çX  ì˙    :  2006/10/16 çáåvÉLÅ[Ç…FCDÇí«â¡
'**
'**********************************************************'
    Dim OutStr          As String   ' CSVèoóÕì‡óe
    Dim KeyTorcd        As String
    Dim KeyFCD          As String                       '2006/10/16í«â¡
    Dim KeySCD          As String
    
    Dim wSuryo          As Currency ' êîó çáåv
    Dim wKingk          As Currency ' ã‡äzçáåv
    Dim wCnt            As Long
    Dim mKin            As Currency ' ã‡äzåvéZÉèÅ[ÉN
'----------------------------------------------------  2005/03/31 Insert
    Dim wwSyosu         As Currency ' è¨êîïî
    Dim wwKingk         As Currency ' êÆêîïî
'----------------------------------------------------  2005/03/31 Insert End
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    App.Title = "îÉä|ã‡ñæç◊ÉäÉXÉg"
    
'''    If Not DataRead Then Exit Sub
    
    '+-------------------------------------------------------- 2005/03/31 Insert
    '+---------- å_ñÒêÊåvéZï˚ñ@éÊìæ
    gSL_Select = " Select * From éÊà¯êÊÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where éÊà¯êÊCD = '" & TKK.å_ñÒêÊ & "'"
    gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = 0"
    If Not TRMRead(gSL_Select, 1, 1) Then
       Call ItemsClearTRM
    End If
    '+-------------------------------------------------------- 2005/03/31 Insert
    
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrt
        .iPrtTitle = "îÉä|ã‡ñæç◊ÉäÉXÉg"
        .iLineMax = 50
        .iLinePitch = 6
        .iPaperSize = p_A3
        .iOrientation = po_Landscape
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
    
    Call PrintProc
    
End Sub
''''*-------------------------*
''''*      àÛç¸ÉfÅ[É^ì«çû
''''*-------------------------*
'''Private Function DataRead()
'''    DataRead = False
'''
'''    gSL_Select = "SELECT ëŒè€îNåé,FCD,éØï éq,å_ñÒêÊ,ïiî‘,éÛì¸,íPâø,SUBSTRING(åüé˚ì˙,1,6) åüé˚ì˙,"
'''    gSL_Select = gSL_Select & " ä«óùî‘çÜ,í˘ê≥óùóR,ÉJÅ[Éh,SUM(êîó )êîó ,SUM(ã‡äz)ã‡äz,"
'''    gSL_Select = gSL_Select & " ISNULL(SY1.ílñºèÃ,'') å_ñÒêÊñº,ISNULL(SY2.ílñºèÃ,'') éØï éqñº"
'''    gSL_Select = gSL_Select & " FROM TEXTîÉä|ã‡ÉeÅ[ÉuÉã TW"
'''    gSL_Select = gSL_Select & " LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY1"
'''    gSL_Select = gSL_Select & "   ON  SY1.ãÊï™ñºèÃ = 'è∆çáëŒè€å_ñÒêÊ'"
'''    gSL_Select = gSL_Select & "   AND SY1.ãÊï™É^ÉCÉv = ''"
'''    gSL_Select = gSL_Select & "   AND SY1.íl = å_ñÒêÊ"
'''    gSL_Select = gSL_Select & " LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY2"
'''    gSL_Select = gSL_Select & "   ON  SY2.ãÊï™ñºèÃ = 'TEXTîÉä|ã‡'"
'''    gSL_Select = gSL_Select & "   AND SY2.ãÊï™É^ÉCÉv = ''"
'''    gSL_Select = gSL_Select & "   AND SY2.íl = å_ñÒêÊ + FCD + éØï éq"
'''    gSL_Select = gSL_Select & PrtSelF(S310_SelCnt)
'''    gSL_Select = gSL_Select & " GROUP BY ëŒè€îNåé,FCD,éØï éq,å_ñÒêÊ,ïiî‘,éÛì¸,SUBSTRING(åüé˚ì˙,1,6),íPâø,"
'''    gSL_Select = gSL_Select & "          ä«óùî‘çÜ,í˘ê≥óùóR,ÉJÅ[Éh,SY1.ílñºèÃ,SY2.ílñºèÃ"
'''    gSL_Select = gSL_Select & " ORDER BY ëŒè€îNåé,FCD,éØï éq,å_ñÒêÊ,ïiî‘,éÛì¸,SUBSTRING(åüé˚ì˙,1,6),íPâø,"
'''    gSL_Select = gSL_Select & "          ä«óùî‘çÜ,í˘ê≥óùóR,ÉJÅ[Éh"
'''
'''    If Not TKKRead(gSL_Select, 0) Then
'''        Call TKKClose
'''        Exit Function
'''    End If
'''    S310_TotalCnt = S310_TotalCnt + CisDB.RecordCount
'''
'''    DataRead = True
'''End Function
'*---------------------*
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
    
    Call CisPrt.ClrLstFld("B*")
    CisPrt.cLine = 99: wCnt = 0
    
    Do Until Not TKK_RDSTS
        With TKK
            If wCnt <> 0 And (KeyFCD <> .FCD Or KeySCD <> .éØï éq) Then                 '2006/10/16ïœçX
'''            If wCnt <> 0 And KeySCD <> .éØï éq Then
                Call TotalWrite(0)  ' éØï éqçáåv
                CisPrt.cLine = 99
            End If
            Call HeadWrite
            
            If RTrim(Mid(.ïiî‘, 11, 2)) <> "" Then
                B1_HINBN = Left(.ïiî‘, 5) & "-" & Mid(.ïiî‘, 6, 5) & "-" & Mid(.ïiî‘, 11, 2)
            Else
                B1_HINBN = Left(.ïiî‘, 5) & "-" & Mid(.ïiî‘, 6, 5)
            End If
            B1_Uke = Trim(.éÛì¸)
            B1_Suryo = Format(Format(.êîó , "#,##0;-#,##0;#"), String(11, "@"))
            B1_TANKA = Format(Format(.íPâø, "#,##0.00;-#,##0.00;#"), String(14, "@"))
    '+-----------------------------------------------  2005/03/31 Insert
            If TRM.åvéZï˚ñ@ = 1 Then    ' åéíPà 
                If Mid(.ïiî‘, 1, 2) = "A5" And .èdó  <> 0 Then
                   wwKingk = Fix(.èdó  * .êîó )
                   wwKingk = wwKingk * .íPâø
                Else
                   wwKingk = .êîó  * .íPâø
                End If
                .ã‡äz = Fix(wwKingk)        'êÿÇËéÃÇƒ
                wwSyosu = wwKingk - .ã‡äz
                
                If TRM.åvéZíPà  = 1 Then 'êÿÇËè„Ç∞
                    If .ã‡äz <> wwKingk Then
                        .ã‡äz = .ã‡äz + 1
                    End If
                End If
                If TRM.åvéZíPà  = 2 Then 'éléÃå‹ì¸
                    If wwSyosu >= 0.5 Then
                        .ã‡äz = .ã‡äz + 1
                    End If
                End If
            End If
    '+-----------------------------------------------  2005/03/31 Insert  End
            
            B1_KINGK = Format(Format(.ã‡äz, "#,##0;-#,##0;#"), String(11, "@"))
            
            B1_YM = Mid(.åüé˚ì˙, 5, 2)
            B1_Kanri = RTrim(.ä«óùî‘çÜ)
            B1_Teisei = RTrim(.í˘ê≥óùóR)
            B1_Card = RTrim(.ÉJÅ[Éh)
    '+----------------------------------------- 2005/03/31 Insert
            If .èdó  <> 0 Then
                B1_Juryo = Format(.èdó , "0.000")
            End If
    '+----------------------------------------- 2005/03/31 Insert End
            
            
            Call CisPrt.PrintFld("B1")
            CisPrt.cLine = CisPrt.cLine + 1
            Call CisPrt.ClrLstFld("B*")
        '*================================*'
        '*  ñæç◊àÛç¸
        '*================================*'
            Call Set_Meisai
        
            wSuryo = wSuryo + .êîó 
            wKingk = wKingk + .ã‡äz
        '--(( KEYëﬁî ))--
            KeyTorcd = .å_ñÒêÊ
            KeyFCD = .FCD                                               '2006/10/16í«â¡
            KeySCD = .éØï éq
            
            With CisPrt
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CisPrt.cDataMaxCnt, "#,###")
            End With
            wCnt = wCnt + 1
        End With

        Call TKKReadNext
    Loop
    Call TKKClose
    
    Call TotalWrite(1)
    
    ' àÛç¸èIóπ
    Call CisPrt.PrintEndDoc

End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite()
'
    HeadWrite = False
    With CisPrt
        ' ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    
        With TKK
            H1_PAGE = CisFun.RSetFld(CisPrt.cPage, 5, "#,###")
            H1_YMD = Format(Now(), "YYYY/MM/DD HH:NN")
            H1_YY = Left(S310_Ym, 4)
            H1_MM = Mid(S310_Ym, 5, 2)
            H1_TORCD = Trim(.å_ñÒêÊ)
            H1_TORNM = Trim(.å_ñÒêÊñº)
            H1_SIKI = Trim(.éØï éq)
            H1_SIKINM = Trim(.éØï éqñº)
        End With
    
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    HeadWrite = True
End Function
'*----------------*
'*  ñæç◊àÛç¸       *
'*----------------*
Private Sub Set_Meisai()
    
    With TKK
    gSL_Select = " SELECT åüé˚ì˙,î[ì¸î‘çÜ,êîó  FROM TEXTîÉä|ã‡ÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & " WHERE å_ñÒêÊ = '" & RTrim(.å_ñÒêÊ) & "'"
    gSL_Select = gSL_Select & "   AND ëŒè€îNåé = '" & RTrim(S310_Ym) & "'"
    gSL_Select = gSL_Select & "   AND FCD = '" & RTrim(.FCD) & "'"
    gSL_Select = gSL_Select & "   AND éØï éq= '" & RTrim(.éØï éq) & "'"
    gSL_Select = gSL_Select & "   AND ïiî‘= '" & RTrim(.ïiî‘) & "'"
    gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(.éÛì¸) & "'"
    gSL_Select = gSL_Select & "   AND SUBSTRING(åüé˚ì˙,1,6) = '" & RTrim(.åüé˚ì˙) & "'"
    gSL_Select = gSL_Select & "   AND íPâø = " & .íPâø & ""
    gSL_Select = gSL_Select & "   AND ä«óùî‘çÜ = '" & RTrim(.ä«óùî‘çÜ) & "'"
    gSL_Select = gSL_Select & "   AND í˘ê≥óùóR = '" & RTrim(.í˘ê≥óùóR) & "'"
    gSL_Select = gSL_Select & "   AND ÉJÅ[Éh = '" & RTrim(.ÉJÅ[Éh) & "'"
    gSL_Select = gSL_Select & " ORDER BY åüé˚ì˙,î[ì¸î‘çÜ,êîó  "
    End With
    gInt = 0
    CisDB.SQL = gSL_Select
    If CisDB.DBRead(0, 1) Then
        Do Until CisDB.Recordset.EOF
            gInt = gInt + 1
            B1_DD(gInt) = Mid(CisDB.Recordset("åüé˚ì˙"), 7, 2)
            B1_NO(gInt) = CisDB.Recordset("î[ì¸î‘çÜ")
            B1_SU(gInt) = Format(Format(CisDB.Recordset("êîó "), "#,##0;-#,##0;#"), String(7, "@"))
            If gInt >= 8 Then
                CisPrt.cLine = CisPrt.cLine + 1
                Call CisPrt.PrintFld("B1")
                Call CisPrt.ClrLstFld("B*")
                Call HeadWrite
                gInt = 0
            End If
            CisDB.Recordset.MoveNext
        Loop
    End If
    CisPrt.cLine = CisPrt.cLine + 1
    If gInt <> 0 Then
        Call CisPrt.PrintFld("B1")
        CisPrt.cLine = CisPrt.cLine + 1
        Call CisPrt.ClrLstFld("B*")
    End If
    CisDB.Recordset.Close
    CisPrt.cLine = CisPrt.cLine + 1

End Sub
'*----------------*
'*   çáåvàÛç¸     *
'*----------------*
Private Sub TotalWrite(Optional KB As Integer = 0)
        
    If KB = 0 Or KB = 1 Then
        If CisPrt.cLine <> 0 Then CisPrt.cLine = CisPrt.cLine + 1
        Call HeadWrite
        
        B1_MIDASI = "Ås  éØï éq çáåvÅ@Åt"
        B1_Suryo = Format(Format(wSuryo, "#,##0;-#,##0;#"), String(11, "@"))
        B1_KINGK = Format(Format(wKingk, "#,##0;-#,##0;#"), String(11, "@"))
        
        Call CisPrt.PrintFld("B1")
        Call CisPrt.ClrLstFld("B*")
    
        S310_TtlSu = S310_TtlSu + wSuryo
        S310_TtlKin = S310_TtlKin + wKingk
        wSuryo = 0: wKingk = 0
    End If
    
    If KB = 1 And S310_GKbn <> "" Then
        If CisPrt.cLine <> 0 Then CisPrt.cLine = CisPrt.cLine + 2
        Call HeadWrite
        If CisPrt.cLine = 0 Then CisPrt.cLine = CisPrt.cLine + 2
        
        B1_MIDASI = "Åy å_ñÒêÊ çáåv Åz"
        B1_Suryo = Format(Format(S310_TtlSu, "#,##0;-#,##0;#"), String(11, "@"))
        B1_KINGK = Format(Format(S310_TtlKin, "#,##0;-#,##0;#"), String(11, "@"))
        
        Call CisPrt.PrintFld("B1")
        Call CisPrt.ClrLstFld("B*")

        S310_TtlSu = 0: S310_TtlKin = 0
    End If
    
    
End Sub

