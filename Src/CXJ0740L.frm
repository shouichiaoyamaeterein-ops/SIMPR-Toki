VERSION 5.00
Begin VB.Form CXJ0740L 
   Caption         =   "èoâ◊ÉeÉLÉXÉgïœä∑ÉäÉXÉg"
   ClientHeight    =   2265
   ClientLeft      =   795
   ClientTop       =   2625
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   16755
   Begin VB.Line B1_K_Line 
      Index           =   19
      X1              =   13950
      X2              =   13950
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   18
      X1              =   13650
      X2              =   13650
      Y1              =   1380
      Y2              =   1650
   End
   Begin VB.Label Label31 
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
      Left            =   13710
      TabIndex        =   57
      Top             =   870
      Width           =   195
   End
   Begin VB.Label Label9 
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
      Left            =   13710
      TabIndex        =   56
      Top             =   1080
      Width           =   195
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
      Index           =   0
      Left            =   13710
      TabIndex        =   55
      Top             =   1425
      Width           =   210
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "î“âÒì˙-ï÷"
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
      Left            =   12270
      TabIndex        =   54
      Top             =   1020
      Width           =   945
   End
   Begin VB.Label B1_BYMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD-ZZ"
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
      TabIndex        =   53
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Line Line4 
      X1              =   15000
      X2              =   16200
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line3 
      X1              =   13650
      X2              =   14850
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "œΩ¿ñ¢ìoò^"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   15240
      TabIndex        =   52
      Top             =   630
      Width           =   810
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "ÉGÉâÅ["
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   14010
      TabIndex        =   51
      Top             =   630
      Width           =   540
   End
   Begin VB.Label B1_NYMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD-ZZ"
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
      Left            =   8550
      TabIndex        =   50
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸ì˙-ï÷"
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
      Left            =   8790
      TabIndex        =   49
      Top             =   990
      Width           =   945
   End
   Begin VB.Line B1_K_Line 
      Index           =   17
      X1              =   15000
      X2              =   15000
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Line B1_Line2 
      X1              =   15000
      X2              =   16200
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   5
      X1              =   15000
      X2              =   15000
      Y1              =   570
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   15000
      X2              =   16200
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Label B1_No 
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
      Left            =   90
      TabIndex        =   48
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "áÇ"
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
      Left            =   150
      TabIndex        =   47
      Top             =   960
      Width           =   240
   End
   Begin VB.Label H1_CVTDATE 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD - ZZZ"
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
      Left            =   2070
      TabIndex        =   46
      Top             =   390
      Width           =   2640
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "ïœä∑èàóùì˙ÅF"
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
      Index           =   2
      Left            =   90
      TabIndex        =   45
      Top             =   390
      Width           =   1980
   End
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
      Left            =   750
      TabIndex        =   44
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
      Left            =   2430
      TabIndex        =   43
      Top             =   960
      Width           =   900
   End
   Begin VB.Label B1_Nouno 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   2445
      TabIndex        =   42
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label B1_FName 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX1"
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
      Left            =   540
      TabIndex        =   41
      Top             =   1470
      Width           =   1800
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "èoâ◊êî"
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
      TabIndex        =   40
      Top             =   990
      Width           =   630
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
      Left            =   10710
      TabIndex        =   39
      Top             =   990
      Width           =   390
   End
   Begin VB.Line Line1 
      X1              =   13650
      X2              =   14850
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Line B1_K_Line 
      Index           =   16
      X1              =   16200
      X2              =   16200
      Y1              =   570
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
      X1              =   14850
      X2              =   14850
      Y1              =   570
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   11
      X1              =   14550
      X2              =   14550
      Y1              =   810
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   10
      X1              =   13650
      X2              =   13650
      Y1              =   570
      Y2              =   1320
   End
   Begin VB.Line B1_K_Line 
      Index           =   9
      X1              =   14250
      X2              =   14250
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
      TabIndex        =   38
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
      TabIndex        =   37
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
      Left            =   14610
      TabIndex        =   36
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
      Left            =   14610
      TabIndex        =   35
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
      Left            =   14310
      TabIndex        =   34
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
      Left            =   14310
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      Left            =   15960
      TabIndex        =   28
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
      Left            =   15960
      TabIndex        =   27
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
      Left            =   15660
      TabIndex        =   26
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
      Left            =   15660
      TabIndex        =   25
      Top             =   870
      Width           =   195
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "íç"
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
      Left            =   14010
      TabIndex        =   24
      Top             =   1080
      Width           =   210
   End
   Begin VB.Label Label4 
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
      Left            =   14010
      TabIndex        =   23
      Top             =   870
      Width           =   210
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
      Left            =   11085
      TabIndex        =   22
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
      Left            =   10740
      TabIndex        =   21
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
      Left            =   13995
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
      Index           =   2
      Left            =   14310
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
      Index           =   3
      Left            =   14610
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
      Index           =   4
      Left            =   15060
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   1425
      Width           =   210
   End
   Begin VB.Line B1_Line 
      X1              =   13650
      X2              =   14850
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   1
      X1              =   14250
      X2              =   14250
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   0
      X1              =   13950
      X2              =   13950
      Y1              =   1365
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   2
      X1              =   14550
      X2              =   14550
      Y1              =   1380
      Y2              =   1635
   End
   Begin VB.Line B1_K_Line 
      Index           =   3
      X1              =   14850
      X2              =   14850
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
      Left            =   8280
      TabIndex        =   15
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
      Left            =   12120
      TabIndex        =   14
      Top             =   195
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
      Left            =   6135
      TabIndex        =   13
      Top             =   945
      Width           =   1155
   End
   Begin VB.Label B1_Tornm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXX"
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
      Left            =   6540
      TabIndex        =   12
      Top             =   1440
      Width           =   1680
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
      Left            =   5760
      TabIndex        =   11
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
      Caption         =   "XXX"
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
      Left            =   5175
      TabIndex        =   10
      Top             =   1440
      Width           =   360
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
      Left            =   13290
      TabIndex        =   9
      Top             =   195
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
      TabIndex        =   8
      Top             =   195
      Width           =   615
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXX"
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
      TabIndex        =   7
      Top             =   1440
      Width           =   1680
   End
   Begin VB.Label Label48 
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
      Left            =   8280
      TabIndex        =   6
      Top             =   960
      Width           =   210
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
      Left            =   9960
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "Å¶Å¶ èoâ◊ÉeÉLÉXÉgïœä∑ÉäÉXÉg Å¶Å¶"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   210
      Width           =   5760
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
      Left            =   5010
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   195
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
      Left            =   3660
      TabIndex        =   1
      Top             =   945
      Width           =   930
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
      Left            =   10080
      TabIndex        =   0
      Top             =   975
      Width           =   630
   End
End
Attribute VB_Name = "CXJ0740L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  èoâ◊ÉeÉLÉXÉgïœä∑ÉäÉXÉg
'**       ÉtÉHÅ[ÉÄID    :  CXJ0740L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2010/03/17
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
    With Text_Prt
        .iPrtTitle = "èoâ◊ÉeÉLÉXÉgïœä∑ÉäÉXÉg"
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
    gCnt1 = 0
    Erase mNo
    
    Do Until Not THN_RDSTS
        With THN
        
            Call Text_Prt.ClrLstFld("B1")
            '*---------------------------------*
            '*  HeadïîÅ@àÛç¸
            '*---------------------------------*
                Call HeadWrite(0)
            
            If BK_FName <> .ÉtÉ@ÉCÉãñº Then
                B1_FName = RTrim(.ÉtÉ@ÉCÉãñº)
                BK_FName = RTrim(.ÉtÉ@ÉCÉãñº)
                
                BK_Nouno = ""
            End If
            
            gCnt1 = gCnt1 + 1
            B1_No = Format(CisFun.RSetFld(gCnt1, 3, "###"))
            
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
            B1_Tornm = RTrim(.î[ì¸êÊñºèÃ)
            B1_Ukeir = RTrim(.éÛì¸)
            If Trim(THN.î[ì¸ì˙) <> "" Then
                With CisFun
                    .DateE_BefVal = THN.î[ì¸ì˙
                    .DateEdit
                    gStr = .DateE_AfVal
                End With
                If Trim(THN.î[ì¸ï÷) <> "" Then
                   gStr = gStr & "-" & THN.î[ì¸ï÷
                Else
                   gStr = gStr & "   "
                End If
                B1_NYMD = gStr
            End If
            B1_Syuyo = CisFun.RSetFld(.é˚óeêî, 7, "#,###")
            B1_Maisu = CisFun.RSetFld(.Ç©ÇÒÇŒÇÒñáêî, 3, "###")
            B1_Suryo = CisFun.RSetFld(.èoâ◊êî, 9, "###")
            If Trim(THN.î“âÒì˙) <> "" Then
                With CisFun
                    .DateE_BefVal = THN.î“âÒì˙
                    .DateEdit
                    gStr = .DateE_AfVal
                End With
                If Trim(THN.î“âÒï÷) <> "" Then
                   gStr = gStr & "-" & THN.î“âÒï÷
                Else
                   gStr = gStr & "   "
                End If
                B1_BYMD = gStr
            End If
            
            If RTrim(.ERRèdï°) <> 0 Then: B1_ERR(0) = "Åõ"
            If RTrim(.ERRéÛíç) <> 0 Then: B1_ERR(1) = "Åõ"
            If RTrim(.ERRé˚óeêî) <> 0 Then: B1_ERR(2) = "Åõ"
            If RTrim(.ERRîwî‘çÜ) <> 0 Then: B1_ERR(3) = "Åõ"
            If RTrim(.ïiî‘âº) <> 0 Then: B1_ERR(4) = "Åõ"
            If RTrim(.ïiéÊâº) <> 0 Then: B1_ERR(5) = "Åõ"
            If RTrim(.éÊà¯êÊâº) <> 0 Then: B1_ERR(6) = "Åõ"
            If RTrim(.éÛì¸âº) <> 0 Then: B1_ERR(7) = "Åõ"
                    
            '*---------------------------------*
            '*  BodyçsÅ@àÛç¸
            '*---------------------------------*
            With Text_Prt
                Call .PrintFld("B1")
                .cLine = .cLine + 1
                Call .ClrLstFld("B1")
            End With
        
        End With
        
        Call THNReadNext
    Loop
    
    Call THNClose
    
    ' àÛç¸èIóπ
    Call Text_Prt.PrintEndDoc
    
End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    
    Dim NewPageFlg As Boolean
'
    HeadWrite = False
    With Text_Prt
        ' ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        NewPageFlg = False
        
        If .iLineMax >= .cLine + 1 And Not NewPageFlg Then Exit Function
        
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    
        H1_PAGE = CisFun.RSetFld(.cPage, 5, "###")
        H1_HYMD = mPrtDate
        
        With CisFun
            .DateE_BefVal = THN.èàóùì˙
            .DateEdit
            gStr = .DateE_AfVal
        End With
        gStr = gStr & " - " & Format(THN.èàóùâÒêî, "000")
        H1_CVTDATE = gStr
        
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    BK_FName = "": BK_Nouno = ""
    HeadWrite = True
End Function
