VERSION 5.00
Begin VB.Form CTOP0010L 
   Caption         =   "TOPPSéÛíçïœä∑ÉGÉâÅ[ÉäÉXÉg"
   ClientHeight    =   2265
   ClientLeft      =   1320
   ClientTop       =   1455
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   16755
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "TEXT-ID"
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
      TabIndex        =   46
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸î‘çÜ"
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
      Left            =   2970
      TabIndex        =   45
      Top             =   960
      Width           =   900
   End
   Begin VB.Label B1_Nouno 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Left            =   3105
      TabIndex        =   44
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label B1_FName 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX1XXXXXXXXX1"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   43
      Top             =   1470
      Width           =   2700
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12120
      TabIndex        =   42
      Top             =   990
      Width           =   585
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "ñáêî"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11280
      TabIndex        =   41
      Top             =   990
      Width           =   390
   End
   Begin VB.Line Line1 
      X1              =   14100
      X2              =   16200
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Line B1_K_Line 
      Index           =   16
      X1              =   16200
      X2              =   16200
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   15
      X1              =   15900
      X2              =   15900
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   14
      X1              =   15600
      X2              =   15600
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   13
      X1              =   15300
      X2              =   15300
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   12
      X1              =   15000
      X2              =   15000
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   11
      X1              =   14700
      X2              =   14700
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   10
      X1              =   14100
      X2              =   14100
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   9
      X1              =   14400
      X2              =   14400
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Label B1_ERR 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   15960
      TabIndex        =   40
      Top             =   1425
      Width           =   210
   End
   Begin VB.Label B1_ERR 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   15660
      TabIndex        =   39
      Top             =   1425
      Width           =   210
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "î‘"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15960
      TabIndex        =   38
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "îw"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15960
      TabIndex        =   37
      Top             =   870
      Width           =   195
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "óe"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15660
      TabIndex        =   36
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "é˚"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15660
      TabIndex        =   35
      Top             =   870
      Width           =   195
   End
   Begin VB.Line B1_K_Line 
      Index           =   8
      X1              =   16200
      X2              =   16200
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   7
      X1              =   15900
      X2              =   15900
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   6
      X1              =   15600
      X2              =   15600
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "éÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15360
      TabIndex        =   34
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "ïi"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15360
      TabIndex        =   33
      Top             =   870
      Width           =   195
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "î‘"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15060
      TabIndex        =   32
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "ïi"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15060
      TabIndex        =   31
      Top             =   870
      Width           =   195
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "éÛ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14760
      TabIndex        =   30
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "éÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14760
      TabIndex        =   29
      Top             =   870
      Width           =   195
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "à¯"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14460
      TabIndex        =   28
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "éÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14460
      TabIndex        =   27
      Top             =   870
      Width           =   195
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "ï°"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14160
      TabIndex        =   26
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "èd"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14160
      TabIndex        =   25
      Top             =   870
      Width           =   195
   End
   Begin VB.Label B1_Hasur 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZZ,ZZ9"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13020
      TabIndex        =   24
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label B1_Suryo 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Z,ZZZ,ZZ9"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11895
      TabIndex        =   23
      Top             =   1455
      Width           =   945
   End
   Begin VB.Label B1_Maisu 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZ9"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11340
      TabIndex        =   22
      Top             =   1455
      Width           =   315
   End
   Begin VB.Label B1_ERR 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   14145
      TabIndex        =   21
      Top             =   1425
      Width           =   210
   End
   Begin VB.Label B1_ERR 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   14460
      TabIndex        =   20
      Top             =   1425
      Width           =   210
   End
   Begin VB.Label B1_ERR 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   14760
      TabIndex        =   19
      Top             =   1425
      Width           =   210
   End
   Begin VB.Label B1_ERR 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   15060
      TabIndex        =   18
      Top             =   1425
      Width           =   210
   End
   Begin VB.Label B1_ERR 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   15360
      TabIndex        =   17
      Top             =   1425
      Width           =   210
   End
   Begin VB.Line B1_Line 
      X1              =   14100
      X2              =   16200
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   1
      X1              =   14400
      X2              =   14400
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   0
      X1              =   14100
      X2              =   14100
      Y1              =   1365
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   2
      X1              =   14700
      X2              =   14700
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   3
      X1              =   15000
      X2              =   15000
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   4
      X1              =   15300
      X2              =   15300
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Label B1_Ukeir 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9750
      TabIndex        =   16
      Top             =   1440
      Width           =   210
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "î≠çsì˙éûÅF"
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
      Index           =   0
      Left            =   11670
      TabIndex        =   15
      Top             =   405
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "î[  ì¸  êÊ"
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
      Index           =   11
      Left            =   7125
      TabIndex        =   14
      Top             =   945
      Width           =   1155
   End
   Begin VB.Label B1_Tornm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7770
      TabIndex        =   13
      Top             =   1440
      Width           =   1350
   End
   Begin VB.Label B1_Torcd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7005
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
   Begin VB.Line H1_K_LINE 
      X1              =   15
      X2              =   16200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
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
      Left            =   5925
      TabIndex        =   11
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label H1_HYMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD hh:ss"
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
      Left            =   12840
      TabIndex        =   10
      Top             =   405
      Width           =   1920
   End
   Begin VB.Label H1_PAGE 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Z,ZZ9"
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
      Left            =   15315
      TabIndex        =   9
      Top             =   405
      Width           =   615
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
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
      Left            =   3810
      TabIndex        =   8
      Top             =   1440
      Width           =   1950
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "éÛ"
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
      Left            =   9690
      TabIndex        =   7
      Top             =   960
      Width           =   240
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZZ,ZZ9"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10380
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "Å¶Å¶ TOPPSéÛíçïœä∑ÉGÉâÅ[ÉäÉXÉg Å¶Å¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   5
      Top             =   285
      Width           =   5565
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "îwî‘çÜ"
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
      Left            =   5940
      TabIndex        =   4
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label28 
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
      Left            =   15960
      TabIndex        =   3
      Top             =   405
      Width           =   225
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ïi    î‘"
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
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      Top             =   945
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "í[êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13275
      TabIndex        =   1
      Top             =   975
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "é˚óeêî"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10470
      TabIndex        =   0
      Top             =   975
      Width           =   630
   End
End
Attribute VB_Name = "CTOP0010L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  TOPPSéÛíçïœä∑ÉGÉâÅ[ÉäÉXÉg
'**       ÉtÉHÅ[ÉÄID    :  CTOP0010L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2004/04/20
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
'#î‰ärëﬁîóp#
    Dim BK_FName        As String
    Dim BK_Nouno        As String
        
    Dim lIx             As Byte         'ìYéö
    Dim TMai            As Long         'ñáêî
    
    Dim mNo(10)         As Long
    
    Dim mPrtDate        As String



'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg   *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd hh:nn")             'î≠çsì˙ÅEéûä‘
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrt
        .iPrtTitle = "TOPPSéÛíçïœä∑ÉGÉâÅ[ÉäÉXÉg"
        .iLineMax = 40
        .iPaperSize = p_A4
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
'*---------------------*
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
    Dim FirstFLG    As Boolean      'èââÒâÒîóp(TRUE:Bodyçs àÛç¸)
    Dim BreakFLG    As Boolean      'KEY-BREAK(FALSE:keyBreak,TRUE:not_KeyBreak)
    
    lIx = 0: TMai = 0
    Erase mNo
    
    gSL_Select = "SELECT TJ.*,"
    gSL_Select = gSL_Select & " ISNULL(TR.ó™èÃ,'') î[ì¸êÊñº,"
    gSL_Select = gSL_Select & " ISNULL(HN.ï\é¶ïiî‘,'') ï\é¶ïiî‘"
    gSL_Select = gSL_Select & " FROM TOPPSéÛíçïœä∑ÉeÉLÉXÉg TJ "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN éÊà¯êÊÉ}ÉXÉ^ TR "
    gSL_Select = gSL_Select & "   ON TR.éÊà¯êÊãÊï™ = 0 "
    gSL_Select = gSL_Select & "  AND TR.éÊà¯êÊCD = TJ.î[ì¸êÊ "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ïiî‘É}ÉXÉ^ HN "
    gSL_Select = gSL_Select & "   ON HN.ïiî‘ = TJ.ïiî‘ "
    gSL_Select = gSL_Select & " WHERE TJ.í[ññî‘çÜ = HOST_NAME() "
    
    If RTrim(gCTOP0010_SaiKB) = "" Then
        gSL_Select = gSL_Select & "   AND (TJ.ERR1 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR2 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR3 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR4 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR5 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR6 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR7 <> '' )"
    Else
        gSL_Select = gSL_Select & "   AND (TJ.ERR2 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR3 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR4 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR5 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR6 <> '' "
        gSL_Select = gSL_Select & "    OR  TJ.ERR7 <> '' )"
    End If
    gSL_Select = gSL_Select & " ORDER BY ÉtÉ@ÉCÉãID,î[ì¸î‘çÜ,î[ïièëçs"
    Call PHWRead(gSL_Select)
    
    
    Do Until Not PHW_RDSTS
        With PHW
        
            Call CisPrt.ClrLstFld("B1")
            '*---------------------------------*
            '*  HeadïîÅ@àÛç¸
            '*---------------------------------*
                Call HeadWrite(0)
            
            If BK_FName <> .ÉtÉ@ÉCÉãID Then
                B1_FName = RTrim(.ÉtÉ@ÉCÉãID)
                BK_FName = RTrim(.ÉtÉ@ÉCÉãID)
                
                BK_Nouno = ""
            End If
            
            If BK_Nouno <> .î[ì¸î‘çÜ Then
                B1_Nouno = RTrim(.î[ì¸î‘çÜ)
                BK_Nouno = RTrim(.î[ì¸î‘çÜ)
            End If
            
            If RTrim(.ï\é¶ïiî‘) <> "" Then
                B1_Hinbn = RTrim(.ï\é¶ïiî‘)
            Else
                B1_Hinbn = RTrim(.ïiî‘)
            End If
            B1_Seban = RTrim(.îwî‘çÜ)
            B1_Torcd = RTrim(.î[ì¸êÊ)
            B1_Tornm = RTrim(.î[ì¸êÊñº)
            B1_Ukeir = RTrim(.éÛì¸)
            B1_Syuyo = CisFun.RSetFld(.é˚óeêî, 7, "#,###")
            B1_Maisu = CisFun.RSetFld(.ñáêî, 3, "###")
            B1_Suryo = CisFun.RSetFld(.î[ì¸êî, 9, "###")
            B1_Hasur = CisFun.RSetFld(.í[êî, 7, "###")
            
            If RTrim(.ERR1) <> "" Then: B1_ERR(1) = "Å~"
            If RTrim(.ERR2) <> "" Then: B1_ERR(2) = "Å~"
            If RTrim(.ERR3) <> "" Then: B1_ERR(3) = "Å~"
            If RTrim(.ERR4) <> "" Then: B1_ERR(4) = "Å~"
            If RTrim(.ERR5) <> "" Then: B1_ERR(5) = "Å~"
            If RTrim(.ERR6) <> "" Then: B1_ERR(6) = "Å~"
            If RTrim(.ERR7) <> "" Then: B1_ERR(7) = "Å~"
                    
            '*---------------------------------*
            '*  BodyçsÅ@àÛç¸
            '*---------------------------------*
            With CisPrt
                Call .PrintFld("B1")
                .cLine = .cLine + 1
                Call .ClrLstFld("B1")
            End With
        
        End With
        
        Call PHWReadNext
    Loop
    
    Call PHWClose
    
    ' àÛç¸èIóπ
    Call CisPrt.PrintEndDoc
    
End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    
    Dim NewPageFlg As Boolean
'
    HeadWrite = False
    With CisPrt
        ' ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        NewPageFlg = False
        
        If .iLineMax >= .cLine + 1 And Not NewPageFlg Then Exit Function
        
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    
        H1_PAGE = CisFun.RSetFld(.cPage, 5, "###")
        H1_HYMD = mPrtDate
        
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    BK_FName = "": BK_Nouno = ""
    HeadWrite = True
End Function
