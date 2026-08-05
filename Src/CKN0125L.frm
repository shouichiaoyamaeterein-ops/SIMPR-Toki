VERSION 5.00
Begin VB.Form CKN0125L 
   Caption         =   "î≠íçÉTÉCÉNÉãàÍóóï\"
   ClientHeight    =   3090
   ClientLeft      =   4140
   ClientTop       =   3480
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   11640
   Begin VB.Line B1_TLine 
      Index           =   0
      X1              =   6195
      X2              =   6195
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      BorderStyle     =   3  'ì_ê¸
      Index           =   1
      X1              =   6840
      X2              =   6840
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      BorderStyle     =   3  'ì_ê¸
      Index           =   2
      X1              =   7485
      X2              =   7485
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      BorderStyle     =   3  'ì_ê¸
      Index           =   3
      X1              =   8115
      X2              =   8115
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      BorderStyle     =   3  'ì_ê¸
      Index           =   8
      X1              =   10635
      X2              =   10635
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      BorderStyle     =   3  'ì_ê¸
      Index           =   9
      X1              =   10005
      X2              =   10005
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      BorderStyle     =   3  'ì_ê¸
      Index           =   10
      X1              =   9375
      X2              =   9375
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      BorderStyle     =   3  'ì_ê¸
      Index           =   11
      X1              =   8745
      X2              =   8745
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      Index           =   12
      X1              =   11235
      X2              =   11235
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      Index           =   17
      X1              =   4740
      X2              =   4740
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      Index           =   18
      X1              =   150
      X2              =   150
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      Index           =   19
      X1              =   3810
      X2              =   3810
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Line B1_TLine 
      Index           =   20
      X1              =   3345
      X2              =   3345
      Y1              =   2415
      Y2              =   2660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   """*""àÛÇÕëOì˙î≠íç"
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
      Index           =   3
      Left            =   9870
      TabIndex        =   69
      Top             =   510
      Width           =   1350
   End
   Begin VB.Line H1_K_LINE 
      Index           =   19
      X1              =   3345
      X2              =   3345
      Y1              =   1890
      Y2              =   2355
   End
   Begin VB.Line H1_K_LINE 
      Index           =   18
      X1              =   3810
      X2              =   3810
      Y1              =   1890
      Y2              =   2355
   End
   Begin VB.Line H1_K_LINE 
      Index           =   17
      X1              =   4740
      X2              =   4740
      Y1              =   1890
      Y2              =   2355
   End
   Begin VB.Line B1_K_LINE 
      BorderStyle     =   3  'ì_ê¸
      Index           =   2
      X1              =   3375
      X2              =   11235
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Label H1_Bikou 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   68
      Top             =   930
      Width           =   4200
   End
   Begin VB.Label H1_EndHYmd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99 - 99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2025
      TabIndex        =   67
      Top             =   1425
      Width           =   2475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ç≈èIî≠íçèàóùì˙-âÒ"
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
      Index           =   17
      Left            =   360
      TabIndex        =   66
      Top             =   1500
      Width           =   1530
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   16
      X1              =   11235
      X2              =   11235
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÇWâÒ"
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
      Index           =   15
      Left            =   10770
      TabIndex        =   65
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label H1_HYKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   10740
      TabIndex        =   64
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label H1_RETime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   10695
      TabIndex        =   63
      Top             =   765
      Width           =   525
   End
   Begin VB.Label B1_Bin 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#9"
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
      Left            =   10830
      TabIndex        =   62
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label H1_HCTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   10695
      TabIndex        =   61
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label H1_HKTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   10695
      TabIndex        =   60
      Top             =   1365
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      Index           =   15
      X1              =   10635
      X2              =   10635
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÇVâÒ"
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
      Index           =   14
      Left            =   10140
      TabIndex        =   59
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label H1_HYKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   10110
      TabIndex        =   58
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label H1_RETime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   10065
      TabIndex        =   57
      Top             =   765
      Width           =   525
   End
   Begin VB.Label B1_Bin 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#9"
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
      Left            =   10200
      TabIndex        =   56
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label H1_HCTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   10065
      TabIndex        =   55
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label H1_HKTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   10065
      TabIndex        =   54
      Top             =   1365
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      Index           =   14
      X1              =   10005
      X2              =   10005
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÇUâÒ"
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
      Index           =   13
      Left            =   9510
      TabIndex        =   53
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label H1_HYKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   9480
      TabIndex        =   52
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label H1_RETime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   9435
      TabIndex        =   51
      Top             =   765
      Width           =   525
   End
   Begin VB.Label B1_Bin 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#9"
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
      Left            =   9570
      TabIndex        =   50
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label H1_HCTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   9435
      TabIndex        =   49
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label H1_HKTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   9435
      TabIndex        =   48
      Top             =   1365
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      Index           =   13
      X1              =   9375
      X2              =   9375
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÇTâÒ"
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
      Index           =   12
      Left            =   8880
      TabIndex        =   47
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label H1_HYKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   8850
      TabIndex        =   46
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label H1_RETime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   8805
      TabIndex        =   45
      Top             =   765
      Width           =   525
   End
   Begin VB.Label B1_Bin 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#9"
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
      Left            =   8940
      TabIndex        =   44
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label H1_HCTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   8805
      TabIndex        =   43
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label H1_HKTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   8805
      TabIndex        =   42
      Top             =   1365
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      Index           =   12
      X1              =   8745
      X2              =   8745
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÇSâÒ"
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
      Index           =   11
      Left            =   8250
      TabIndex        =   41
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label H1_HYKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   8220
      TabIndex        =   40
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label H1_RETime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   8175
      TabIndex        =   39
      Top             =   765
      Width           =   525
   End
   Begin VB.Label B1_Bin 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#9"
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
      Left            =   8310
      TabIndex        =   38
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label H1_HCTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   8175
      TabIndex        =   37
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label H1_HKTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   8175
      TabIndex        =   36
      Top             =   1365
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      Index           =   11
      X1              =   8115
      X2              =   8115
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÇRâÒ"
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
      Index           =   7
      Left            =   7620
      TabIndex        =   35
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label H1_HYKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   7590
      TabIndex        =   34
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label H1_RETime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   7545
      TabIndex        =   33
      Top             =   765
      Width           =   525
   End
   Begin VB.Label B1_Bin 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#9"
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
      Left            =   7680
      TabIndex        =   32
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label H1_HCTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   7545
      TabIndex        =   31
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label H1_HKTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   7545
      TabIndex        =   30
      Top             =   1365
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      Index           =   10
      X1              =   7485
      X2              =   7485
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÇQâÒ"
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
      Index           =   4
      Left            =   6990
      TabIndex        =   29
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label H1_HYKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6960
      TabIndex        =   28
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label H1_RETime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6915
      TabIndex        =   27
      Top             =   765
      Width           =   525
   End
   Begin VB.Label B1_Bin 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#9"
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
      Left            =   7050
      TabIndex        =   26
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label H1_HCTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6915
      TabIndex        =   25
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label H1_HKTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6915
      TabIndex        =   24
      Top             =   1365
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      Index           =   9
      X1              =   5175
      X2              =   5175
      Y1              =   705
      Y2              =   1905
   End
   Begin VB.Line H1_K_LINE 
      Index           =   8
      X1              =   150
      X2              =   150
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Line H1_K_LINE 
      Index           =   7
      X1              =   5175
      X2              =   11235
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Line H1_K_LINE 
      Index           =   6
      X1              =   5175
      X2              =   11235
      Y1              =   1305
      Y2              =   1305
   End
   Begin VB.Line H1_K_LINE 
      Index           =   5
      X1              =   5175
      X2              =   11235
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Line H1_K_LINE 
      Index           =   4
      X1              =   6840
      X2              =   6840
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Line H1_K_LINE 
      Index           =   3
      X1              =   6195
      X2              =   6195
      Y1              =   705
      Y2              =   2355
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   2
      X1              =   150
      X2              =   11235
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ç≈èIî≠íçì˙-ï÷"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4815
      TabIndex        =   23
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Line T1_K_LINE 
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   11235
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line B1_K_LINE 
      BorderStyle     =   3  'ì_ê¸
      Index           =   3
      X1              =   3840
      X2              =   11235
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line B1_K_LINE 
      Index           =   1
      X1              =   150
      X2              =   11235
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Label H1_HKTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6270
      TabIndex        =   21
      Top             =   1365
      Width           =   525
   End
   Begin VB.Label H1_HCTime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6270
      TabIndex        =   20
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "à¯éÊéûä‘"
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
      Index           =   10
      Left            =   5325
      TabIndex        =   19
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "âÒé˚éûä‘"
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
      Index           =   9
      Left            =   5325
      TabIndex        =   18
      Top             =   795
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "î≠íçéûä‘"
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
      Index           =   8
      Left            =   5325
      TabIndex        =   17
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label H1_Page 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10740
      TabIndex        =   16
      Top             =   195
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "ï≈"
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
      Index           =   2
      Left            =   11100
      TabIndex        =   15
      Top             =   210
      Width           =   225
   End
   Begin VB.Label B1_Cycle 
      AutoSize        =   -1  'True
      Caption         =   "9-99-99"
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
      Left            =   3855
      TabIndex        =   14
      Top             =   2430
      Width           =   840
   End
   Begin VB.Label B1_Bin 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#9"
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
      Left            =   6405
      TabIndex        =   13
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label H1_RETime 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6270
      TabIndex        =   12
      Top             =   765
      Width           =   525
   End
   Begin VB.Label H1_HYKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6315
      TabIndex        =   11
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label B1_Tornm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   975
      TabIndex        =   10
      Top             =   2430
      Width           =   2400
   End
   Begin VB.Label B1_Torcd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   195
      TabIndex        =   9
      Top             =   2430
      Width           =   840
   End
   Begin VB.Label B1_Ukeir 
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
      Left            =   3450
      TabIndex        =   8
      Top             =   2430
      Width           =   240
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   1
      X1              =   150
      X2              =   11235
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "éÛì¸"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3390
      TabIndex        =   7
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÇPâÒ"
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
      Index           =   6
      Left            =   6345
      TabIndex        =   6
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "íãñÈãŒ"
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
      Index           =   5
      Left            =   5325
      TabIndex        =   5
      Top             =   1665
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÉTÉCÉNÉã"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3885
      TabIndex        =   4
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "édÅ@ì¸Å@êÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1050
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   11235
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Label Label1 
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
      Height          =   180
      Index           =   1
      Left            =   8160
      TabIndex        =   2
      Top             =   225
      Width           =   540
   End
   Begin VB.Label H1_HYmd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99 hh:mm"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8790
      TabIndex        =   1
      Top             =   195
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Å¶Å@î≠íçÉeÅ[ÉuÉãàÍóóï\Å@Å¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   285
      Width           =   3900
   End
   Begin VB.Label B1_ENDHDate 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "YY/MM/DD-99"
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
      Left            =   4815
      TabIndex        =   22
      Top             =   2430
      Width           =   1320
   End
End
Attribute VB_Name = "CKN0125L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  î≠íçÉeÅ[ÉuÉãàÍóóï\Å@î≠çs
'**       ÉtÉHÅ[ÉÄID    :  CKN0125L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2005/03/03
'**       ïœ  çX  ì˙    :  2007/06/04  àÍóóï\ÇÃç≈èIçsãyÇ—ÅAéüï≈êÊì™çsÇ™Ç®Ç©ÇµÇ¢
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim mPrtDate        As String
    Dim mPrtTime        As String
'   àÛéöÉTÉvÉåÉXóp
    Dim Key_Torcd       As String
    Dim Key_Cycle       As String
    Dim Key_Ukeir       As String
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")          'î≠çsì˙
    mPrtTime = Format(Now(), "hh:nn")               'î≠çséûä‘
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrtCKN0125
        .iPrtTitle = "î≠íçÉeÅ[ÉuÉãàÍóóï\"
        .iLineMax = 56
        .iLinePitch = 6
        .iPaperSize = p_A4
       '.iOrientation = po_PORTRAIT

        If Not .PrintIni Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(1) = "     Å@ÉvÉäÉìÉ^èâä˙ê›íËÉGÉâÅ[        "
                .MB_Title = "àÛç¸¥◊∞"
                .MB_Button = Error
                .MBOX
            End With
            Exit Sub
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

    Call CisPrtCKN0125.ClrLstFld("B1")

'*---------------------------------*
'*  î≠íçÉTÉCÉNÉãÉ}ÉXÉ^ì«çûÇ›
'*---------------------------------*
    Do Until Not HCM_RDSTS

        Call HeadWrite(0)

        '*---------------------------------*
        '*  BodyçsÅ@ï“èW
        '*---------------------------------*
        With HCM
            If Key_Torcd <> RTrim(.édì¸êÊ) Or _
                Key_Ukeir <> RTrim(.éÛì¸) Or _
                Key_Cycle <> RTrim(.ÉTÉCÉNÉã) Then
                If RTrim(Key_Torcd) <> "" Then
                    '*---------------------------------*
                    '*  BodyïîÅ@àÛç¸
                    '*---------------------------------*
                    With CisPrtCKN0125
                        .PrintFld ("B1*")
                        .cLine = .cLine + 1
                        .ClrLstFld ("B1")
                    End With
                    B1_K_LINE(1).Visible = False
                    B1_K_LINE(2).Visible = False
                End If
                
                If Key_Torcd <> RTrim(.édì¸êÊ) Then
                    B1_Torcd = RTrim(.édì¸êÊ)
                    B1_Tornm = RTrim(.édì¸êÊñº)
                    B1_K_LINE(1).Visible = True
                Else
                End If
                If Key_Torcd <> RTrim(.édì¸êÊ) Or Key_Ukeir <> RTrim(.éÛì¸) Then
                    B1_Ukeir = RTrim(.éÛì¸)
                    B1_K_LINE(2).Visible = True
                End If
                If HCM.ójì˙î≠íç = 0 Then
                    B1_Cycle = RTrim(.ï“èWÉTÉCÉNÉã)
                Else
                    B1_Cycle = EditWeekDay
                End If
                
                If Trim(.ç≈èIî≠íçì˙) <> "" Then
                   With CisFun
                        .DateE_BefVal = HCM.ç≈èIî≠íçì˙
                        .DateEdit
                        B1_ENDHDate = Mid(.DateE_AfVal, 3)
                   End With
                End If
                If RTrim(.ç≈èIî≠íçï÷) <> "" Then
                    B1_ENDHDate = B1_ENDHDate & "-" & RTrim(.ç≈èIî≠íçï÷)
                    
                End If
            End If
            
            If .î≠íçèàóùâÒêî <> 0 Then
                B1_Bin(.î≠íçèàóùâÒêî) = CisFun.Val2(.êUèoï÷)
                If .î≠çsâ¡éZì˙ <> 0 Then
                    B1_Bin(.î≠íçèàóùâÒêî) = B1_Bin(.î≠íçèàóùâÒêî) & "*"
                End If
            End If
        
            With CisPrtCKN0125
            '   ÉJÉEÉìÉ^Å@ÇtÇo
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
            End With
        '   àÛéöÉTÉvÉåÉXóp
            Key_Torcd = RTrim(.édì¸êÊ)
            Key_Ukeir = RTrim(.éÛì¸)
            Key_Cycle = RTrim(.ÉTÉCÉNÉã)
        End With

        Call HCMReadNext(1)

    '   TailïîÅ@àÛç¸
       'With CisPrtCKN0125
       '    If Not HCM_RDSTS Or .iLineMax = .cLine Then
       '        .PrintFld ("T1*")
       '    End If
       'End With
    Loop
    Call HCMClose(1)
    If RTrim(Key_Torcd) <> "" Then
        '*---------------------------------*
        '*  BodyïîÅ@àÛç¸
        '*---------------------------------*
        With CisPrtCKN0125
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
            .PrintFld ("T1*")
        End With
    End If

'   àÛç¸èIóπ
    Call CisPrtCKN0125.PrintEndDoc
End Function
'*---------------------------------------------------------------*
'*      ÉwÉbÉ_ïîàÛç¸                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False

    With CisPrtCKN0125
    '   ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function
        If .cPage > 0 Then

'---   ---   ---   ---   ---   ---   ---< 2007/06/04 ADD >
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
'---   ---   ---   ---   ---   ---   ---< 2007/06/04 ADD END >

            .PrintFld ("T1*")
        End If

        .cLine = 0
        .cPage = .cPage + 1
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
        gSL_Select = " SELECT * FROM î≠íçÉTÉCÉNÉãê›íËÉ}ÉXÉ^ "
        gSL_Select = gSL_Select & " ORDER BY âÒêî"
        If HCSRead(gSL_Select, 0, 0) Then
            Do Until Not HCS_RDSTS
                If HCS.âÒêî = 0 Then
                    H1_Bikou = HCS.îıçl
                    If RTrim(HCS.ç≈èIî≠íçì˙) <> "" Then
                        H1_EndHYmd = Left(HCS.ç≈èIî≠íçì˙, 4) & "/" & Mid(HCS.ç≈èIî≠íçì˙, 5, 2) & "/" & Mid(HCS.ç≈èIî≠íçì˙, 7, 2)
                        H1_EndHYmd = H1_EndHYmd & " - " & HCS.ç≈èIâÒ
                    End If
                Else
                    If RTrim(HCS.âÒé˚éûä‘) <> "" And RTrim(HCS.âÒé˚éûä‘) <> "0000" Then
                        H1_RETime(HCS.âÒêî) = Left(HCS.âÒé˚éûä‘, 2) & ":" & Mid(HCS.âÒé˚éûä‘, 3, 2)
                    End If
                    If RTrim(HCS.î≠íçéûä‘) <> "" Then
                        H1_HCTime(HCS.âÒêî) = Left(HCS.î≠íçéûä‘, 2) & ":" & Mid(HCS.î≠íçéûä‘, 3, 2)
                    End If
                    If RTrim(HCS.à¯éÊéûä‘) <> "" And RTrim(HCS.à¯éÊéûä‘) <> "0000" Then
                        H1_HKTime(HCS.âÒêî) = Left(HCS.à¯éÊéûä‘, 2) & ":" & Mid(HCS.à¯éÊéûä‘, 3, 2)
                    End If
                    If HCS.íãñÈãÊï™ = 0 Then
                        H1_HYKbn(HCS.âÒêî) = "íããŒ"
                    Else
                        H1_HYKbn(HCS.âÒêî) = "ñÈãŒ"
                    End If
                End If
            
                Call HCSReadNext
            Loop
        End If
        Call HCSClose(0)
    '   àÛéöÉTÉvÉåÉXóp
        Key_Torcd = ""
        Key_Ukeir = ""
        Key_Cycle = ""

    '   î≠çsì˙ÅEï≈
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")

        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With

    HeadWrite = True
End Function
'==========================================================='
'   ójì˙î≠íçÅ@ï\é¶
'==========================================================='
Private Function EditWeekDay() As String
    EditWeekDay = ""
    gSL_Select = "SELECT * FROM éÊà¯êÊéÛì¸É}ÉXÉ^"
    gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(HCM.édì¸êÊ) & "'"
    gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(HCM.éÛì¸) & "'"
    gSL_Select = gSL_Select & "   AND éÊà¯êÊãÊï™ = 1"
    If Not TUMRead(gSL_Select, 1, 0) Then Exit Function
    If TUM.ójì˙1 <> 0 Then EditWeekDay = EditWeekDay & "åé"
    If TUM.ójì˙2 <> 0 Then EditWeekDay = EditWeekDay & "âŒ"
    If TUM.ójì˙3 <> 0 Then EditWeekDay = EditWeekDay & "êÖ"
    If TUM.ójì˙4 <> 0 Then EditWeekDay = EditWeekDay & "ñÿ"
    If TUM.ójì˙5 <> 0 Then EditWeekDay = EditWeekDay & "ã‡"

End Function

