VERSION 5.00
Object = "{0D300FC0-B2EA-11D1-8D3B-444553540000}#1.30#0"; "Qrmaker.ocx"
Begin VB.Form CXHZ0360L 
   Caption         =   "端数ラベル"
   ClientHeight    =   5940
   ClientLeft      =   1515
   ClientTop       =   5220
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   11640
   Begin QRMAKERLib.QRmaker B1_QRmaker1 
      Height          =   1080
      Left            =   9255
      TabIndex        =   19
      Top             =   1080
      Width           =   1185
      _Version        =   65566
      _ExtentX        =   2090
      _ExtentY        =   1905
      _StockProps     =   1
      Picture         =   "CXHZ0360L.frx":0000
   End
   Begin VB.Label B1_Zaisy 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3450
      TabIndex        =   40
      Top             =   3480
      Width           =   1800
   End
   Begin VB.Label B1_TaniM 
      AutoSize        =   -1  'True
      Caption         =   "単位"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8265
      TabIndex        =   39
      Top             =   4035
      Width           =   390
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   765
      Index           =   11
      Left            =   8160
      Top             =   3990
      Width           =   735
   End
   Begin VB.Label B1_ZaisyM 
      AutoSize        =   -1  'True
      Caption         =   "材種"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2775
      TabIndex        =   38
      Top             =   3495
      Width           =   390
   End
   Begin VB.Label B1_SebanM 
      AutoSize        =   -1  'True
      Caption         =   "背番号"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   37
      Top             =   4035
      Width           =   585
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   765
      Index           =   8
      Left            =   2700
      Top             =   3990
      Width           =   1695
   End
   Begin VB.Label B1_Sunpo 
      AutoSize        =   -1  'True
      Caption         =   "9999.99x9999.99x9999.99"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3270
      TabIndex        =   36
      Top             =   1980
      Width           =   5520
   End
   Begin VB.Label B1_Tani 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8190
      TabIndex        =   35
      Top             =   4320
      Width           =   660
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   525
      Index           =   6
      Left            =   2700
      Top             =   2700
      Width           =   6210
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   6
      Height          =   810
      Index           =   5
      Left            =   2700
      Top             =   1785
      Width           =   6195
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   6
      Height          =   810
      Index           =   4
      Left            =   2700
      Top             =   990
      Width           =   6195
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   765
      Index           =   3
      Left            =   6075
      Top             =   3990
      Width           =   2025
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   765
      Index           =   2
      Left            =   4440
      Top             =   3990
      Width           =   1590
   End
   Begin VB.Label B1_ZainmM 
      AutoSize        =   -1  'True
      Caption         =   "材料名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2775
      TabIndex        =   34
      Top             =   2850
      Width           =   585
   End
   Begin VB.Label B1_Zainm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3540
      TabIndex        =   33
      Top             =   2835
      Width           =   4800
   End
   Begin VB.Label B1_SunpoM 
      AutoSize        =   -1  'True
      Caption         =   "寸法"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2775
      TabIndex        =   32
      Top             =   2100
      Width           =   390
   End
   Begin VB.Label B1_HinbnM 
      AutoSize        =   -1  'True
      Caption         =   "材質"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2775
      TabIndex        =   31
      Top             =   1305
      Width           =   390
   End
   Begin VB.Label B1_Zaist 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "XXXXXXXXX1XXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   27.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3420
      TabIndex        =   30
      Top             =   1125
      Width           =   5100
   End
   Begin VB.Label B1_SyuyoM 
      AutoSize        =   -1  'True
      Caption         =   "収容数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6105
      TabIndex        =   29
      Top             =   4035
      Width           =   585
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "99,999.999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      TabIndex        =   28
      Top             =   4290
      Width           =   1950
   End
   Begin VB.Label B1_SyukiM 
      AutoSize        =   -1  'True
      Caption         =   "収容器"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4500
      TabIndex        =   27
      Top             =   4035
      Width           =   585
   End
   Begin VB.Label B1_Syuki 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4455
      TabIndex        =   26
      Top             =   4320
      Width           =   1515
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   660
      Index           =   0
      Left            =   2700
      Top             =   3285
      Width           =   2730
   End
   Begin VB.Label B1_Seban 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2775
      TabIndex        =   25
      Top             =   4290
      Width           =   1575
   End
   Begin VB.Label B1_Zaist25 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   24
      Top             =   1215
      Width           =   5625
   End
   Begin VB.Label B1_Hasu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999,999.999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5985
      TabIndex        =   23
      Top             =   3390
      Width           =   2640
   End
   Begin VB.Label B1_HasuM 
      AutoSize        =   -1  'True
      Caption         =   "端数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5535
      TabIndex        =   22
      Top             =   3495
      Width           =   390
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   6
      Height          =   645
      Index           =   10
      Left            =   5460
      Top             =   3285
      Width           =   3435
   End
   Begin VB.Label B1_Okiba 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   21
      Top             =   4320
      Width           =   1365
   End
   Begin VB.Label B1_OkibaM 
      AutoSize        =   -1  'True
      Caption         =   "置場"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   540
      TabIndex        =   20
      Top             =   4350
      Width           =   390
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   585
      Index           =   9
      Left            =   480
      Top             =   4170
      Width           =   2145
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   1350
      Index           =   1
      Left            =   9315
      Top             =   2955
      Width           =   1380
   End
   Begin VB.Label B1_HYmdM 
      AutoSize        =   -1  'True
      Caption         =   "発行日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   9195
      TabIndex        =   18
      Top             =   4575
      Width           =   540
   End
   Begin VB.Label B1_HYmd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9825
      TabIndex        =   17
      Top             =   4545
      Width           =   1050
   End
   Begin VB.Label B1_UkeirM 
      AutoSize        =   -1  'True
      Caption         =   "受入"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      Top             =   3075
      Width           =   390
   End
   Begin VB.Label B1_Ukeir 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   36
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9585
      TabIndex        =   15
      Top             =   3345
      Width           =   900
   End
   Begin VB.Label B1_NBin 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   540
      TabIndex        =   14
      Top             =   3300
      Width           =   600
   End
   Begin VB.Label B1_NYmdM 
      AutoSize        =   -1  'True
      Caption         =   "納入日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   255
      TabIndex        =   13
      Top             =   2700
      Width           =   540
   End
   Begin VB.Label B1_NYmd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   2940
      Width           =   2100
   End
   Begin VB.Label B1_JisyaKj 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8700
      TabIndex        =   11
      Top             =   435
      Width           =   2400
   End
   Begin VB.Label B1_NonyuKj 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   10
      Top             =   1830
      Width           =   2400
   End
   Begin VB.Label B1_Nonyu 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   990
      TabIndex        =   9
      Top             =   1335
      Width           =   630
   End
   Begin VB.Label B1_NonyuNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   2400
   End
   Begin VB.Label B1_NonyuM 
      AutoSize        =   -1  'True
      Caption         =   "手配商社"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   225
      TabIndex        =   7
      Top             =   1335
      Width           =   720
   End
   Begin VB.Label B1_SiireKj 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   705
      Width           =   2400
   End
   Begin VB.Label B1_SiireM 
      AutoSize        =   -1  'True
      Caption         =   "仕入先"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   225
      TabIndex        =   5
      Top             =   195
      Width           =   540
   End
   Begin VB.Label B1_Tehai 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3030
      TabIndex        =   4
      Top             =   330
      Width           =   720
   End
   Begin VB.Label B1_SiireNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   435
      Width           =   2400
   End
   Begin VB.Label B1_Siire 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   975
      TabIndex        =   2
      Top             =   195
      Width           =   630
   End
   Begin VB.Label B1_JisyaNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8685
      TabIndex        =   1
      Top             =   150
      Width           =   2400
   End
   Begin VB.Label B1_PrintKb 
      AutoSize        =   -1  'True
      Caption         =   "端数ラベル"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   26.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4410
      TabIndex        =   0
      Top             =   165
      Width           =   2700
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      Height          =   4890
      Index           =   7
      Left            =   30
      Top             =   0
      Width           =   11220
   End
End
Attribute VB_Name = "CXHZ0360L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  [材料]端数ラベル　発行
'**       フォームID    :  CXHZ0360L
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/24
'**       変  更  日    :  2008/04/11 By CIS - 背番号８桁対応
'**
'**********************************************************'
'#比較退避用#
    Dim MaxGyo          As Byte         '頁行数
    Dim GMaisu          As Long         '現枚数
    Dim TMaisu          As Long         '総枚数
    
    Dim mPrtDate        As String
    Dim mPrtTime        As String
    Dim mQRData         As String * 161
    Dim mJisyaNm        As String
    Dim mJisyaKj        As String
    Dim mHMoto          As String

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             '発行日
    mPrtTime = Format(Now(), "hh:nn")                  '発行時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCXHZ0360
        .iPrtTitle = "[材料]端数ラベル"
        .iLineMax = 80
'        .iLinePitch = 8
        .iPaperSize = p_A4
        .iOrientation = po_Portrait
        
        If Not .PrintIni Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(1) = "     　プリンタ初期設定エラー        "
                .MB_Title = "印刷ｴﾗｰ"
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
'*　 << 印  刷  処  理 >>
'*---------------------------------------------------------------*
Private Function PrintProc()
    Dim GGyo    As Integer
' システム情報(発行元)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    Call IRNRead("", 1)
    mHMoto = Mid(IRN.oyian, 1, 7)
'*---------------------------------*
'*  自社名称・工場名を求める
'*---------------------------------*
' システム情報(自社工場)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    Call IRNRead("", 1)
' 取引先マスタより名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & Left(IRN.oyian, 1) & "'"
    If TRMRead(gSL_Select, 1) Then
        mJisyaNm = TRM.略称
        mJisyaKj = TRM.工場名
    End If
    
'*---------------------------------*
'*  初期設定
'*---------------------------------*
    MaxGyo = 3: GGyo = 0
    Call CisPrtCXHZ0360.ClrLstFld("B1")
    Call HeadWrite(0)
    
'*---------------------------------*
'*  材料発注テーブル読込み
'*---------------------------------*
    Do Until Not ZCT_RDSTS
        With ZCT
        '*---------------------------------*
        '*  Body行　印刷（枚数分繰返し）
        '*---------------------------------*
            If GGyo + 1 > MaxGyo Then
                Call HeadWrite(0)
                GGyo = 0
            End If
            GGyo = GGyo + 1
                
            B1_PrintKb = "端数ラベル"
            B1_Tehai = RTrim(.手配区分名)
            B1_JisyaNm = RTrim(mJisyaNm)
            B1_JisyaKj = RTrim(mJisyaKj)
            
            B1_SiireM = "仕入先"
            B1_Siire = RTrim(.仕入先CD)
            If RTrim(.仕入先工場CD) <> "" Then
                B1_Siire = B1_Siire & "-" & RTrim(.仕入先工場CD)
            End If
            B1_SiireNm = RTrim(.仕入先名)
            B1_SiireKj = RTrim(.仕入先工場名)
            If RTrim(.納入先) <> "" Then
                If RTrim(.手配区分) <> "1" Then
                    B1_NonyuM = "手配商社"
                Else
                    B1_NonyuM = "納入先"
                End If
                B1_Nonyu = RTrim(.納入先CD)
                If RTrim(.納入先工場CD) <> "" Then
                    B1_Nonyu = B1_Nonyu & "-" & RTrim(.納入先工場CD)
                End If
                B1_NonyuNm = RTrim(.納入先名)
                B1_NonyuKj = RTrim(.納入先工場名)
            End If
            
            B1_NYmdM = "納入日"
            CisFun.DateE_BefVal = .納入日
            CisFun.DateEdit (J_YMD)
            B1_NYmd = CisFun.DateE_AfVal
            If RTrim(.納入便) <> "" Then
                B1_NBin = RTrim(.納入便) & "便"
            End If
            B1_OkibaM = "置場"
            B1_Okiba = RTrim(.置場)
            
            B1_HinbnM = "材質"
            If Len(RTrim(.材質)) <= 16 Then
                B1_Zaist = RTrim(.材質)
            Else
                B1_Zaist25 = RTrim(.材質)
            End If
            B1_SunpoM = "寸法"
            B1_Sunpo = RTrim(.表示寸法)
            B1_ZainmM = "材料名"
            B1_Zainm = RTrim(.材料名)
            B1_ZaisyM = "材種"
            If RTrim(.材種コード) = "" Then
                B1_Zaisy = RTrim(.材種名)
            Else
                B1_Zaisy = RTrim(.材種コード)
            End If
            
            B1_TaniM = "単位"
            B1_Tani = RTrim(.手配単位名)
            B1_SebanM = "背番号"
            B1_Seban = RTrim(.背番号)
            B1_SyukiM = "収容器"
            B1_Syuki = RTrim(.収容器)
            B1_SyuyoM = "収容数"
            If .収容数 <> 0 Then
                B1_Syuyo = Format(.収容数, "#,##0.000")
            End If
            
            B1_HasuM = "端数"
            B1_Hasu = Format(.端数, "#,##0.000")
            
            B1_UkeirM = "受入"
            B1_Ukeir = RTrim(.受入)
            B1_HYmdM = "発行日"
            B1_HYmd = RTrim(mPrtDate)
        'QR
            Call QR_Edit
            B1_QRmaker1.InputData = mQRData
            B1_QRmaker1.Refresh
        'BODY部印刷
            Call CisPrtCXHZ0360.PrintFld("B1*")
            CisPrtCXHZ0360.cLine = CisPrtCXHZ0360.cLine + 23
            Call CisPrtCXHZ0360.ClrLstFld("B1")
            
            '*---------------------------------*
            '*  カウンタ　ＵＰ
            '*---------------------------------*
            With CisPrtCXHZ0360
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXHZ0360_PrtCnt, "#,###")
            End With
            
        End With
        Call ZCTReadNext(1)
    Loop
    
    Call ZCTClose(1)
    
    ' 印刷終了
    Call CisPrtCXHZ0360.PrintEndDoc
    
End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim NewPageFlg As Boolean
    HeadWrite = False
    
    With CisPrtCXHZ0360
        NewPageFlg = False
        .cLine = 0
        .cPage = .cPage + 1
'
        .ClrLstFld ("H1")   'HEAD部内容消去
'
        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With
    HeadWrite = True
End Function
'*---------------------------------------------------------------*
'*   << QR内容編集 >>                                             *
'*---------------------------------------------------------------*
Private Sub QR_Edit()
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
    'mQRData = Space(161)
    mQRData = Space(164)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    
    Mid(mQRData, 1, 7) = Left(mHMoto, 7)                '発行元
    Mid(mQRData, 8, 5) = gCompany                       '自社コード
    Mid(mQRData, 13, 1) = "H"                           'QR識別(R:臨時かんばん、S:現品票、H:端数ラベル)
    Mid(mQRData, 14, 1) = "Z"                           '品目(0:受注、1:発注、Z:材料)
    Mid(mQRData, 15, 25) = ZCT.材質                     '材質
    Mid(mQRData, 40, 7) = Format(ZCT.板厚, "0000.00")   '収容数
    Mid(mQRData, 47, 7) = Format(ZCT.幅, "0000.00")     '幅
    Mid(mQRData, 54, 7) = Format(ZCT.長さ, "0000.00")   '長さ
    Mid(mQRData, 61, 7) = ZCT.仕入先                    '取引先コード
    Mid(mQRData, 68, 2) = ZCT.受入                      '取引先受入
    Mid(mQRData, 70, 1) = Format(ZCT.手配区分, "0")     '手配区分
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
'    Mid(mQRData, 71, 5) = ZCT.背番号                    '背番号
'    Mid(mQRData, 76, 5) = ZCT.納入番号                  '納入番号
'    Mid(mQRData, 81, 8) = ZCT.納入日                    '納入日
'    Mid(mQRData, 89, 2) = ZCT.納入便                    '納入便
'    Mid(mQRData, 91, 7) = ZCT.納入先                    '納入先コード
'    Mid(mQRData, 98, 2) = ZCT.納入先受入                '納入先受入
'    Mid(mQRData, 100, 10) = Format(ZCT.収容数, "000000.000")  '収容数
'    Mid(mQRData, 110, 10) = ZCT.収容器                   '収容器
'    Mid(mQRData, 120, 12) = Format(ZCT.端数, "00000000.000")  '数量(1枚分)
'    Mid(mQRData, 132, 3) = "001"                        '現枚数
'    Mid(mQRData, 135, 3) = "001"                        '総枚数
'    CisFun.DateE_AfVal = mPrtDate
'    CisFun.DateEditUn
'    Mid(mQRData, 138, 8) = CisFun.DateE_BefVal          '発行日
'    Mid(mQRData, 146, 5) = mPrtTime                     '発行時間
'    Mid(mQRData, 151, 1) = Format(ZCT.発注区分, "0")    '発注区分
'    Mid(mQRData, 152, 10) = ZCT.置場                    '置場
    Mid(mQRData, 71, 8) = ZCT.背番号                    '背番号
    Mid(mQRData, 79, 5) = ZCT.納入番号                  '納入番号
    Mid(mQRData, 84, 8) = ZCT.納入日                    '納入日
    Mid(mQRData, 92, 2) = ZCT.納入便                    '納入便
    Mid(mQRData, 94, 7) = ZCT.納入先                    '納入先コード
    Mid(mQRData, 101, 2) = ZCT.納入先受入                '納入先受入
    Mid(mQRData, 103, 10) = Format(ZCT.収容数, "000000.000")  '収容数
    Mid(mQRData, 113, 10) = ZCT.収容器                   '収容器
    Mid(mQRData, 123, 12) = Format(ZCT.端数, "00000000.000")  '数量(1枚分)
    Mid(mQRData, 135, 3) = "001"                        '現枚数
    Mid(mQRData, 138, 3) = "001"                        '総枚数
    CisFun.DateE_AfVal = mPrtDate
    CisFun.DateEditUn
    Mid(mQRData, 141, 8) = CisFun.DateE_BefVal          '発行日
    Mid(mQRData, 149, 5) = mPrtTime                     '発行時間
    Mid(mQRData, 154, 1) = Format(ZCT.発注区分, "0")    '発注区分
    Mid(mQRData, 155, 10) = ZCT.置場                    '置場
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    
End Sub

