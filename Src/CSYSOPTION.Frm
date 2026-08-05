VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CSYSOPTION 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "SIMPR-A 環境設定"
   ClientHeight    =   10875
   ClientLeft      =   -450
   ClientTop       =   285
   ClientWidth     =   15420
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10875
   ScaleWidth      =   15420
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   1545
      Left            =   8130
      ScaleHeight     =   1485
      ScaleWidth      =   6960
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   7695
      Width           =   7020
      Begin Cis3D_v60.CIS3D lb_KPrinter 
         Height          =   315
         Left            =   1875
         Top             =   105
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   556
         BackColor       =   16777152
         Caption         =   ""
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         cPositionX      =   60
      End
      Begin VB.CheckBox H1Ck_SPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "なし"
         Height          =   195
         Left            =   3840
         TabIndex        =   60
         Top             =   495
         Width           =   705
      End
      Begin CisText_V60.CisText iB1_SHeigh 
         Height          =   285
         Left            =   5685
         TabIndex        =   61
         Top             =   435
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         cFormat         =   "##"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "##"
         MaxLength       =   2
         cDataType       =   1
         cDataReplace    =   1
         cCaps           =   2
         cFbComma        =   0
         cFdAutoFormat   =   1
         cGFormat        =   "##"
         cILength        =   2
      End
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   750
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1323
         BackColor       =   16711680
         ForeColor       =   16777215
         Caption         =   "かんばん設    定"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         cPositionY      =   -100
      End
      Begin Cis3D_v60.CIS3D PB_PrtSet 
         Height          =   300
         Index           =   0
         Left            =   6570
         Top             =   105
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         ForeColor       =   0
         Caption         =   "…"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   8
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   735
         Index           =   5
         Left            =   0
         Top             =   750
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         BackColor       =   16711680
         ForeColor       =   16777215
         Caption         =   "　 自給　　チケット"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         cPositionX      =   45
         cPositionY      =   -100
      End
      Begin Cis3D_v60.CIS3D lb_CPrinter 
         Height          =   315
         Left            =   1875
         Top             =   990
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   556
         BackColor       =   16777152
         Caption         =   ""
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D PB_PrtSet 
         Height          =   300
         Index           =   1
         Left            =   6570
         Top             =   990
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         ForeColor       =   0
         Caption         =   "…"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   8
         cButton         =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "印刷先"
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
         Index           =   1
         Left            =   1275
         TabIndex        =   110
         Top             =   1050
         Width           =   630
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   1215
         X2              =   6975
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   1215
         X2              =   6975
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '透明
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   6435
         TabIndex        =   109
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '透明
         Caption         =   "＋"
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
         Left            =   5445
         TabIndex        =   107
         Top             =   480
         Width           =   210
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '透明
         Caption         =   "高さ調整"
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
         Left            =   4605
         TabIndex        =   106
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "印刷先"
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
         Index           =   0
         Left            =   1275
         TabIndex        =   105
         Top             =   165
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '透明
         Caption         =   "( 仕切かんばん発行"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2130
         TabIndex        =   104
         Top             =   480
         Width           =   1650
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '透明
         Caption         =   "㎜"
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
         Left            =   6045
         TabIndex        =   103
         Top             =   540
         Width           =   195
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D12 
      Height          =   570
      Left            =   8145
      Top             =   9240
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   1005
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "SIMPR-A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Index           =   0
         Left            =   2010
         TabIndex        =   97
         Top             =   -90
         Width           =   2130
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "SIMPR-A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   0
         Left            =   2085
         TabIndex        =   98
         Top             =   -30
         Width           =   1995
      End
   End
   Begin VB.PictureBox B1_Area2 
      Height          =   6990
      Left            =   8130
      ScaleHeight     =   6930
      ScaleWidth      =   6960
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   675
      Width           =   7020
      Begin VB.PictureBox Picture1 
         Height          =   6540
         Left            =   4035
         ScaleHeight     =   6480
         ScaleWidth      =   2895
         TabIndex        =   87
         Top             =   420
         Width           =   2955
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   6465
            Left            =   0
            Top             =   0
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   11404
            BackColor       =   14737632
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            Begin Cis3D_v60.CIS3D ConnectionStr 
               Height          =   1005
               Index           =   6
               Left            =   30
               Top             =   5415
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   1773
               BackColor       =   14737632
               Caption         =   "QRﾎｯﾊﾟﾘｰﾀﾞ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cBoderStyle     =   1
               cPositionX      =   50
               cPositionY      =   60
               Begin VB.OptionButton H1Op_KiKi7 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "使用"
                  Height          =   255
                  Index           =   0
                  Left            =   900
                  TabIndex        =   80
                  Top             =   435
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton H1Op_KiKi7 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "未使用"
                  Height          =   255
                  Index           =   1
                  Left            =   1695
                  TabIndex        =   81
                  Top             =   435
                  Width           =   855
               End
               Begin VB.CommandButton PB_Set 
                  Caption         =   "設定"
                  Height          =   285
                  Index           =   6
                  Left            =   255
                  TabIndex        =   79
                  TabStop         =   0   'False
                  Top             =   405
                  Width           =   570
               End
               Begin VB.Label Connection 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  '透明
                  Caption         =   "TX-2400"
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
                  Index           =   6
                  Left            =   1905
                  TabIndex        =   88
                  Top             =   120
                  Width           =   840
               End
               Begin VB.Shape Shape7 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   1  '不透明
                  Height          =   300
                  Left            =   1860
                  Shape           =   4  '丸みのある長方形
                  Top             =   60
                  Width           =   930
               End
            End
            Begin Cis3D_v60.CIS3D ConnectionStr 
               Height          =   1005
               Index           =   5
               Left            =   30
               Top             =   4410
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   1773
               BackColor       =   14737632
               Caption         =   "かんばん手差ﾘｰﾀﾞｰ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cBoderStyle     =   1
               cPositionX      =   50
               cPositionY      =   60
               Begin VB.OptionButton H1Op_KiKi6 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "未使用"
                  Height          =   255
                  Index           =   1
                  Left            =   1695
                  TabIndex        =   78
                  Top             =   435
                  Width           =   855
               End
               Begin VB.OptionButton H1Op_KiKi6 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "使用"
                  Height          =   255
                  Index           =   0
                  Left            =   900
                  TabIndex        =   77
                  Top             =   435
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.CommandButton PB_Set 
                  Caption         =   "設定"
                  Height          =   285
                  Index           =   5
                  Left            =   255
                  TabIndex        =   76
                  TabStop         =   0   'False
                  Top             =   405
                  Width           =   570
               End
               Begin VB.Label Connection 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  '透明
                  Caption         =   "TX-1301"
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
                  Index           =   5
                  Left            =   1935
                  TabIndex        =   89
                  Top             =   105
                  Width           =   735
               End
               Begin VB.Shape Shape6 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   1  '不透明
                  Height          =   300
                  Left            =   1860
                  Shape           =   4  '丸みのある長方形
                  Top             =   45
                  Width           =   930
               End
            End
            Begin Cis3D_v60.CIS3D ConnectionStr 
               Height          =   1005
               Index           =   3
               Left            =   30
               Top             =   2415
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   1773
               BackColor       =   14737632
               Caption         =   "QRﾊﾝﾄﾞｽｷｬﾅ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cBoderStyle     =   1
               cPositionX      =   50
               cPositionY      =   60
               Begin VB.CommandButton PB_Set 
                  Caption         =   "設定"
                  Height          =   285
                  Index           =   3
                  Left            =   255
                  TabIndex        =   71
                  TabStop         =   0   'False
                  Top             =   405
                  Width           =   570
               End
               Begin VB.OptionButton H1Op_KiKi4 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "未使用"
                  Height          =   255
                  Index           =   1
                  Left            =   1695
                  TabIndex        =   73
                  Top             =   435
                  Width           =   855
               End
               Begin VB.OptionButton H1Op_KiKi4 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "使用"
                  Height          =   255
                  Index           =   0
                  Left            =   900
                  TabIndex        =   72
                  Top             =   435
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.Label Connection 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  '透明
                  Caption         =   "GT10Q-SR"
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
                  Index           =   3
                  Left            =   1905
                  TabIndex        =   90
                  Top             =   120
                  Width           =   840
               End
               Begin VB.Shape Shape4 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   1  '不透明
                  Height          =   300
                  Left            =   1860
                  Shape           =   4  '丸みのある長方形
                  Top             =   60
                  Width           =   930
               End
            End
            Begin Cis3D_v60.CIS3D ConnectionStr 
               Height          =   690
               Index           =   2
               Left            =   1935
               Top             =   1815
               Visible         =   0   'False
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   1217
               BackColor       =   14737632
               Caption         =   "かんばんﾌﾟﾘﾝﾀ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cBoderStyle     =   1
               cPositionX      =   50
               cPositionY      =   60
               Begin VB.CommandButton PB_Set 
                  Caption         =   "設定"
                  Height          =   285
                  Index           =   2
                  Left            =   255
                  TabIndex        =   68
                  TabStop         =   0   'False
                  Top             =   345
                  Visible         =   0   'False
                  Width           =   570
               End
               Begin VB.OptionButton H1Op_KiKi3 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "未使用"
                  Height          =   255
                  Index           =   1
                  Left            =   1695
                  TabIndex        =   70
                  Top             =   375
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.OptionButton H1Op_KiKi3 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "使用"
                  Height          =   255
                  Index           =   0
                  Left            =   900
                  TabIndex        =   69
                  Top             =   375
                  Value           =   -1  'True
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.Label Connection 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  '透明
                  Caption         =   "B-858"
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
                  Index           =   2
                  Left            =   1935
                  TabIndex        =   91
                  Top             =   90
                  Width           =   735
               End
               Begin VB.Shape Shape3 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   1  '不透明
                  Height          =   300
                  Left            =   1860
                  Shape           =   4  '丸みのある長方形
                  Top             =   30
                  Width           =   930
               End
            End
            Begin Cis3D_v60.CIS3D PFXLabel1 
               Height          =   405
               Index           =   6
               Left            =   0
               Top             =   -15
               Width           =   2850
               _ExtentX        =   5027
               _ExtentY        =   714
               BackColor       =   4210752
               ForeColor       =   16777215
               Caption         =   "《 接続機器関連 》"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   14.25
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
            Begin Cis3D_v60.CIS3D ConnectionStr 
               Height          =   990
               Index           =   4
               Left            =   30
               Top             =   3420
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   1746
               BackColor       =   14737632
               Caption         =   "ﾊﾞｰｺｰﾄﾞ(ｷｰﾎﾞｰﾄﾞｲﾝﾀｰﾌｪｲｽ）"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cBoderStyle     =   1
               cPositionX      =   50
               cPositionY      =   60
               Begin VB.OptionButton H1Op_KiKi5 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "未使用"
                  Height          =   255
                  Index           =   1
                  Left            =   1695
                  TabIndex        =   75
                  Top             =   420
                  Width           =   855
               End
               Begin VB.OptionButton H1Op_KiKi5 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "使用"
                  Height          =   255
                  Index           =   0
                  Left            =   900
                  TabIndex        =   74
                  Top             =   420
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.Label Connection 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  '透明
                  Caption         =   "ｷｰﾎﾞｰﾄﾞｲﾝﾀｰﾌｪｲｽ"
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
                  Index           =   4
                  Left            =   885
                  TabIndex        =   92
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   1710
               End
            End
            Begin Cis3D_v60.CIS3D ConnectionStr 
               Height          =   1005
               Index           =   0
               Left            =   30
               Top             =   405
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   1773
               BackColor       =   14737632
               Caption         =   "かんばんｿｰﾀｰ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cBoderStyle     =   1
               cPositionX      =   50
               cPositionY      =   60
               Begin VB.CommandButton PB_Set 
                  Caption         =   "設定"
                  Height          =   285
                  Index           =   0
                  Left            =   255
                  TabIndex        =   62
                  TabStop         =   0   'False
                  Top             =   405
                  Width           =   570
               End
               Begin VB.OptionButton H1Op_KiKi1 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "未使用"
                  Height          =   255
                  Index           =   1
                  Left            =   1695
                  TabIndex        =   64
                  Top             =   435
                  Width           =   855
               End
               Begin VB.OptionButton H1Op_KiKi1 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "使用"
                  Height          =   255
                  Index           =   0
                  Left            =   900
                  TabIndex        =   63
                  Top             =   435
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.Label Connection 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  '透明
                  Caption         =   "TX-597G"
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
                  Index           =   0
                  Left            =   1935
                  TabIndex        =   93
                  Top             =   105
                  Width           =   735
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   1  '不透明
                  Height          =   300
                  Left            =   1860
                  Shape           =   4  '丸みのある長方形
                  Top             =   45
                  Width           =   930
               End
            End
            Begin Cis3D_v60.CIS3D ConnectionStr 
               Height          =   1005
               Index           =   1
               Left            =   30
               Top             =   1410
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   1773
               BackColor       =   14737632
               Caption         =   "かんばん手差ﾘｰﾀﾞｰ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cBoderStyle     =   1
               cPositionX      =   50
               cPositionY      =   60
               Begin VB.CommandButton PB_Set 
                  Caption         =   "設定"
                  Height          =   285
                  Index           =   1
                  Left            =   255
                  TabIndex        =   65
                  TabStop         =   0   'False
                  Top             =   405
                  Width           =   570
               End
               Begin VB.OptionButton H1Op_KiKi2 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "使用"
                  Height          =   255
                  Index           =   0
                  Left            =   900
                  TabIndex        =   66
                  Top             =   435
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton H1Op_KiKi2 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "未使用"
                  Height          =   255
                  Index           =   1
                  Left            =   1695
                  TabIndex        =   67
                  Top             =   435
                  Width           =   855
               End
               Begin VB.Label Connection 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  '透明
                  Caption         =   "TX-1609"
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
                  Index           =   1
                  Left            =   1935
                  TabIndex        =   94
                  Top             =   105
                  Width           =   735
               End
               Begin VB.Shape Shape2 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   1  '不透明
                  Height          =   300
                  Left            =   1860
                  Shape           =   4  '丸みのある長方形
                  Top             =   45
                  Width           =   930
               End
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   6480
         Left            =   30
         Top             =   435
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   11430
         BackColor       =   14737632
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   840
            Left            =   0
            Top             =   3900
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1482
            BackColor       =   14737632
            Caption         =   "イメージ格納先"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iB1_ImageDir 
               Height          =   330
               Left            =   60
               TabIndex        =   57
               Top             =   405
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CSYSOPTION.frx":0000
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_ImageDir 
               Height          =   300
               Left            =   3585
               Top             =   405
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   529
               ForeColor       =   0
               Caption         =   "…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   8
               cButton         =   -1  'True
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   960
            Left            =   0
            Top             =   2940
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1693
            BackColor       =   14737632
            Caption         =   "TEXT変換元格納先"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   30
            Begin VB.CheckBox H1Ck_Jisya 
               BackColor       =   &H00E0E0E0&
               Caption         =   "無"
               Height          =   240
               Left            =   3480
               TabIndex        =   56
               Top             =   675
               Width           =   450
            End
            Begin CisText_V60.CisText iB1_TextDir 
               Height          =   345
               Left            =   60
               TabIndex        =   54
               Top             =   270
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CSYSOPTION.frx":00FE
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_TextDir 
               Height          =   315
               Left            =   3570
               Top             =   255
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               ForeColor       =   0
               Caption         =   "…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   8
               cButton         =   -1  'True
            End
            Begin CisText_V60.CisText iB1_LogDel 
               Height          =   345
               Left            =   1710
               TabIndex        =   55
               Top             =   600
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   609
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cCaps           =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   3
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "自社ｺｰﾄﾞﾁｪｯｸ"
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
               Left            =   2175
               TabIndex        =   96
               Top             =   675
               Width           =   1260
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "変換ログ保持日数"
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
               Left            =   15
               TabIndex        =   95
               Top             =   660
               Width           =   1680
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   840
            Left            =   0
            Top             =   2100
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1482
            BackColor       =   14737632
            Caption         =   "平準化月超過ファイル出力先"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iB1_HeiFDir 
               Height          =   345
               Left            =   60
               TabIndex        =   53
               Top             =   405
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CSYSOPTION.frx":01FC
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_HeiFDir 
               Height          =   315
               Left            =   3585
               Top             =   405
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               ForeColor       =   0
               Caption         =   "…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   8
               cButton         =   -1  'True
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   840
            Left            =   0
            Top             =   1260
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1482
            BackColor       =   14737632
            Caption         =   "JAMA伝票ファイル出力先"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iB1_JamaFDir 
               Height          =   345
               Left            =   60
               TabIndex        =   52
               Top             =   405
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CSYSOPTION.frx":02FA
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_JamaFDir 
               Height          =   315
               Left            =   3585
               Top             =   405
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               ForeColor       =   0
               Caption         =   "…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   8
               cButton         =   -1  'True
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   405
            Index           =   2
            Left            =   -15
            Top             =   0
            Width           =   3990
            _ExtentX        =   7038
            _ExtentY        =   714
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《 フォルダ指定 》"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   840
            Left            =   0
            Top             =   420
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1482
            BackColor       =   14737632
            Caption         =   "自社伝票ファイル出力先"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iB1_JDenFDir 
               Height          =   345
               Left            =   60
               TabIndex        =   51
               Top             =   405
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CSYSOPTION.frx":03F8
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_JDenFDir 
               Height          =   315
               Left            =   3570
               Top             =   405
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               ForeColor       =   0
               Caption         =   "…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   8
               cButton         =   -1  'True
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   840
            Left            =   0
            Top             =   5580
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1482
            BackColor       =   14737632
            Caption         =   "支給チケットファイル出力先"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iB1_STicketFDir 
               Height          =   345
               Left            =   60
               TabIndex        =   59
               Top             =   405
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CSYSOPTION.frx":04F6
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_STicketFDir 
               Height          =   315
               Left            =   3585
               Top             =   405
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               ForeColor       =   0
               Caption         =   "…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   8
               cButton         =   -1  'True
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D17 
            Height          =   840
            Left            =   0
            Top             =   4740
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1482
            BackColor       =   14737632
            Caption         =   "自給チケットファイル出力先"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iB1_JTicketFDir 
               Height          =   345
               Left            =   60
               TabIndex        =   58
               Top             =   405
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CSYSOPTION.frx":05F4
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_JTicketFDir 
               Height          =   315
               Left            =   3570
               Top             =   405
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               ForeColor       =   0
               Caption         =   "…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   8
               cButton         =   -1  'True
            End
         End
      End
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   450
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   794
         BackColor       =   16711680
         ForeColor       =   16777215
         Caption         =   "【 クライアント設定 】"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
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
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9945
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   13605
      ScaleHeight     =   555
      ScaleWidth      =   1515
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   30
      Width           =   1575
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   555
         Left            =   750
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnType         =   1
         BtnCaption      =   "終了"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711680
         BtnShadowColor  =   16777215
         PFCaption       =   "( F12 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   12648447
         PFShadowColor   =   0
         PFBackColor     =   16711680
      End
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   555
         Left            =   0
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnCaption      =   "更新"
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
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10035
      Top             =   195
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   9150
      Left            =   30
      ScaleHeight     =   9090
      ScaleWidth      =   8010
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   675
      Width           =   8070
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   1470
         Left            =   0
         Top             =   420
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   2593
         BackColor       =   14737632
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D CIS3D19 
            Height          =   480
            Left            =   2835
            Top             =   1170
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "受注画面入力 初期値"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   200
            Begin VB.OptionButton H1Op_JGamen 
               BackColor       =   &H00E0E0E0&
               Caption         =   "一納番"
               Height          =   255
               Index           =   0
               Left            =   2835
               TabIndex        =   114
               Top             =   135
               Value           =   -1  'True
               Width           =   1050
            End
            Begin VB.OptionButton H1Op_JGamen 
               BackColor       =   &H00E0E0E0&
               Caption         =   "多納番"
               Height          =   255
               Index           =   1
               Left            =   3930
               TabIndex        =   113
               Top             =   135
               Width           =   1170
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   405
            Index           =   8
            Left            =   0
            Top             =   0
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   714
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《 受注関連 》"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
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
         Begin Cis3D_v60.CIS3D CIS3D14 
            Height          =   960
            Index           =   0
            Left            =   30
            Top             =   435
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   1693
            BackColor       =   14737632
            Caption         =   "売上処理"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   200
            Begin VB.OptionButton H1Op_Uriage 
               BackColor       =   &H00E0E0E0&
               Caption         =   "検収データ"
               Height          =   255
               Index           =   0
               Left            =   1260
               TabIndex        =   0
               Top             =   195
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton H1Op_Uriage 
               BackColor       =   &H00E0E0E0&
               Caption         =   "出荷データ"
               Height          =   255
               Index           =   1
               Left            =   1260
               TabIndex        =   1
               Top             =   555
               Width           =   1335
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D3 
            Height          =   480
            Index           =   2
            Left            =   2835
            Top             =   435
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "受注かんばん読取ｴﾗｰﾘｽﾄ 発行"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_JKErrPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行なし"
               Height          =   255
               Index           =   1
               Left            =   3915
               TabIndex        =   3
               Top             =   135
               Width           =   1170
            End
            Begin VB.OptionButton H1Op_JKErrPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行有"
               Height          =   255
               Index           =   0
               Left            =   2835
               TabIndex        =   2
               Top             =   135
               Value           =   -1  'True
               Width           =   1050
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D14 
            Height          =   480
            Index           =   1
            Left            =   2835
            Top             =   915
            Visible         =   0   'False
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "TOPPS空出荷指示書 発行"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   200
            Begin VB.OptionButton H1Op_JToppsPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行なし"
               Height          =   255
               Index           =   1
               Left            =   3930
               TabIndex        =   5
               Top             =   135
               Width           =   1170
            End
            Begin VB.OptionButton H1Op_JToppsPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行有"
               Height          =   255
               Index           =   0
               Left            =   2835
               TabIndex        =   4
               Top             =   135
               Value           =   -1  'True
               Width           =   1050
            End
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   7200
         Left            =   0
         Top             =   1875
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   12700
         BackColor       =   14737632
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   1305
            Index           =   5
            Left            =   30
            Top             =   3750
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   2302
            BackColor       =   14737632
            Caption         =   "自社伝票"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   100
            cPositionY      =   250
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   855
            Index           =   4
            Left            =   435
            Top             =   4215
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   1508
            BackColor       =   14737632
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.CheckBox H1Ck_JDate 
               BackColor       =   &H00E0E0E0&
               Caption         =   "受領日印字あり"
               Height          =   240
               Left            =   915
               TabIndex        =   25
               Top             =   480
               Width           =   1575
            End
            Begin VB.CheckBox H1Ck_JHanko 
               BackColor       =   &H00E0E0E0&
               Caption         =   "印字なし"
               Height          =   240
               Left            =   915
               TabIndex        =   24
               Top             =   90
               Width           =   1500
            End
            Begin CisText_V60.CisText iB1_JUe 
               Height          =   375
               Left            =   2985
               TabIndex        =   26
               Top             =   75
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXX"
               MaxLength       =   8
               IMEMode         =   4
               cCaps           =   2
            End
            Begin CisText_V60.CisText iB1_JShita 
               Height          =   360
               Left            =   2985
               TabIndex        =   27
               Top             =   435
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXX"
               MaxLength       =   8
               IMEMode         =   4
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   840
               Index           =   6
               Left            =   0
               Top             =   0
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   1482
               BackColor       =   14737632
               Caption         =   "受領書受領印"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cBoderStyle     =   1
               cPositionX      =   100
               cPositionY      =   250
            End
            Begin VB.Label JShita 
               BackColor       =   &H00E0E0E0&
               Caption         =   "下段"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   2475
               TabIndex        =   100
               Top             =   510
               Width           =   525
            End
            Begin VB.Label JUe 
               BackColor       =   &H00E0E0E0&
               Caption         =   "上段"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2475
               TabIndex        =   99
               Top             =   105
               Width           =   525
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   450
            Index           =   2
            Left            =   30
            Top             =   1350
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   794
            BackColor       =   14737632
            Caption         =   "かんばん読取り 即時発注"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_KPool 
               BackColor       =   &H00E0E0E0&
               Caption         =   "即時"
               Height          =   255
               Index           =   1
               Left            =   3510
               TabIndex        =   12
               Top             =   105
               Width           =   765
            End
            Begin VB.OptionButton H1Op_KPool 
               BackColor       =   &H00E0E0E0&
               Caption         =   "指定"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   11
               Top             =   105
               Value           =   -1  'True
               Width           =   795
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   480
            Index           =   1
            Left            =   30
            Top             =   3255
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "平準化月超過発注"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.CheckBox H1Ck_KHei 
               BackColor       =   &H00E0E0E0&
               Caption         =   "ＣＳＶ"
               Height          =   240
               Left            =   1755
               TabIndex        =   19
               Top             =   120
               Width           =   840
            End
            Begin VB.OptionButton H1Op_KHei 
               BackColor       =   &H00E0E0E0&
               Caption         =   "回収"
               Height          =   255
               Index           =   0
               Left            =   2625
               TabIndex        =   20
               Top             =   98
               Value           =   -1  'True
               Width           =   720
            End
            Begin VB.OptionButton H1Op_KHei 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発注"
               Height          =   255
               Index           =   1
               Left            =   3510
               TabIndex        =   21
               Top             =   98
               Width           =   810
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D100 
            Height          =   480
            Index           =   1
            Left            =   30
            Top             =   420
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "発行元"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   50
            cPositionY      =   120
            Begin CisText_V60.CisText iB1_HMoto 
               Height          =   360
               Left            =   750
               TabIndex        =   6
               Top             =   90
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXX"
               MaxLength       =   7
               IMEMode         =   2
               cCaps           =   2
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D3 
            Height          =   720
            Index           =   0
            Left            =   4575
            Top             =   900
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   1270
            BackColor       =   14737632
            Caption         =   "臨時かんばん発行"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_RinPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行なし"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   32
               Top             =   375
               Width           =   1170
            End
            Begin VB.OptionButton H1Op_RinPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行有"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   31
               Top             =   105
               Value           =   -1  'True
               Width           =   1050
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   405
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   7980
            _ExtentX        =   14076
            _ExtentY        =   714
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《 発注関連 》"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
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
         Begin Cis3D_v60.CIS3D UC_3D4 
            Height          =   720
            Index           =   0
            Left            =   4575
            Top             =   1635
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   1270
            BackColor       =   14737632
            Caption         =   "現品票発行"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_GenPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行なし"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   34
               Top             =   375
               Width           =   1080
            End
            Begin VB.OptionButton H1Op_GenPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行有"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   33
               Top             =   105
               Value           =   -1  'True
               Width           =   1080
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   720
            Index           =   0
            Left            =   4575
            Top             =   2370
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   1270
            BackColor       =   14737632
            Caption         =   "端数ラベル発行"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_HasPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行なし"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   36
               Top             =   375
               Width           =   1080
            End
            Begin VB.OptionButton H1Op_HasPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行有"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   35
               Top             =   105
               Value           =   -1  'True
               Width           =   1080
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D100 
            Height          =   480
            Index           =   0
            Left            =   435
            Top             =   3735
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   50
            cPositionY      =   105
            Begin CisText_V60.CisText iB1_Kojyo 
               Height          =   345
               Left            =   1650
               TabIndex        =   22
               Top             =   60
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "X"
               MaxLength       =   1
               cCaps           =   2
            End
            Begin CisText_V60.CisText iB1_Kiban 
               Height          =   345
               Left            =   3165
               TabIndex        =   23
               Top             =   60
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XX"
               MaxLength       =   2
               IMEMode         =   2
               cCaps           =   2
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "工場"
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
               Left            =   1065
               TabIndex        =   101
               Top             =   120
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "機番"
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
               Left            =   2580
               TabIndex        =   85
               Top             =   135
               Width           =   390
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D100 
            Height          =   480
            Index           =   2
            Left            =   30
            Top             =   5070
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "【JAMA伝票】  工場"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   50
            cPositionY      =   105
            Begin CisText_V60.CisText iB1_JKojyo 
               Height          =   345
               Left            =   2040
               TabIndex        =   28
               Top             =   60
               Width           =   225
               _ExtentX        =   397
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "X"
               MaxLength       =   1
               cCaps           =   2
            End
            Begin CisText_V60.CisText iB1_JKiban 
               Height          =   345
               Left            =   3570
               TabIndex        =   29
               Top             =   75
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XX"
               MaxLength       =   2
               IMEMode         =   2
               cCaps           =   2
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "機番"
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
               Left            =   2925
               TabIndex        =   86
               Top             =   150
               Width           =   390
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   570
            Index           =   1
            Left            =   0
            Top             =   6015
            Width           =   7965
            _ExtentX        =   14049
            _ExtentY        =   1005
            BackColor       =   8421504
            ForeColor       =   16777215
            Caption         =   "《材料発注管理》"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cPositionX      =   120
            Begin VB.CheckBox H1chk_Reg 
               BackColor       =   &H00808080&
               Caption         =   "登録解除"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF80&
               Height          =   225
               Left            =   2340
               TabIndex        =   45
               Top             =   180
               Width           =   1335
            End
            Begin Cis3D_v60.CIS3D Zai_Pnl 
               Height          =   495
               Index           =   0
               Left            =   3735
               Top             =   45
               Width           =   4185
               _ExtentX        =   7382
               _ExtentY        =   873
               BackColor       =   14737632
               Caption         =   "内示対象データ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   50
               Begin VB.OptionButton H1Op_ZNJYMKB 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "当月"
                  Height          =   225
                  Index           =   0
                  Left            =   1590
                  TabIndex        =   46
                  Top             =   165
                  Value           =   -1  'True
                  Width           =   690
               End
               Begin VB.OptionButton H1Op_ZNJYMKB 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "翌月"
                  Height          =   225
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   47
                  Top             =   165
                  Width           =   690
               End
               Begin VB.OptionButton H1Op_ZNJYMKB 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "翌々月"
                  Height          =   225
                  Index           =   2
                  Left            =   3210
                  TabIndex        =   48
                  Top             =   165
                  Width           =   870
               End
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   1350
            Index           =   3
            Left            =   4575
            Top             =   3120
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   2381
            BackColor       =   14737632
            Caption         =   "　発注明細"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   100
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   660
            Left            =   4980
            Top             =   3810
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   1164
            BackColor       =   14737632
            Caption         =   "【空明細】"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   100
            Begin VB.OptionButton H1Op_KarPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行なし"
               Height          =   255
               Index           =   0
               Left            =   1515
               TabIndex        =   40
               Top             =   330
               Width           =   1080
            End
            Begin VB.OptionButton H1Op_KarPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行有"
               Height          =   255
               Index           =   1
               Left            =   1515
               TabIndex        =   39
               Top             =   90
               Value           =   -1  'True
               Width           =   1080
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D11 
            Height          =   660
            Left            =   4980
            Top             =   3120
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   1164
            BackColor       =   14737632
            Caption         =   "【  控  】"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   100
            Begin VB.OptionButton H1Op_HIkPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行有"
               Height          =   255
               Index           =   1
               Left            =   1515
               TabIndex        =   37
               Top             =   90
               Value           =   -1  'True
               Width           =   1080
            End
            Begin VB.OptionButton H1Op_HIkPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行なし"
               Height          =   255
               Index           =   0
               Left            =   1515
               TabIndex        =   38
               Top             =   345
               Width           =   1080
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D100 
            Height          =   450
            Index           =   3
            Left            =   30
            Top             =   5565
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   794
            BackColor       =   14737632
            Caption         =   "【支給チケット】  "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   50
            cPositionY      =   105
            Begin CisText_V60.CisText iB1_CKiban 
               Height          =   345
               Left            =   3570
               TabIndex        =   30
               Top             =   75
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XX"
               MaxLength       =   2
               IMEMode         =   2
               cCaps           =   2
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "機番"
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
               Left            =   2925
               TabIndex        =   108
               Top             =   150
               Width           =   390
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   480
            Index           =   7
            Left            =   30
            Top             =   2310
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "納入日･便算出"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_NYmdBin 
               BackColor       =   &H00E0E0E0&
               Caption         =   "時間"
               Height          =   255
               Index           =   0
               Left            =   1980
               TabIndex        =   15
               Top             =   105
               Value           =   -1  'True
               Width           =   795
            End
            Begin VB.OptionButton H1Op_NYmdBin 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発注ﾃｰﾌﾞﾙ"
               Height          =   255
               Index           =   1
               Left            =   3165
               TabIndex        =   16
               Top             =   105
               Width           =   1185
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   480
            Index           =   8
            Left            =   30
            Top             =   2790
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "完納予定日・便算出"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_KYmd 
               BackColor       =   &H00E0E0E0&
               Caption         =   "自動算出"
               Height          =   255
               Index           =   1
               Left            =   3165
               TabIndex        =   18
               Top             =   105
               Width           =   1095
            End
            Begin VB.OptionButton H1Op_KYmd 
               BackColor       =   &H00E0E0E0&
               Caption         =   "手動入力"
               Height          =   255
               Index           =   0
               Left            =   1980
               TabIndex        =   17
               Top             =   105
               Value           =   -1  'True
               Width           =   1065
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   480
            Index           =   9
            Left            =   1830
            Top             =   420
            Width           =   6120
            _ExtentX        =   10795
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "かんばん流動履歴日数       かんばん滞留期間(稼動日)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin CisText_V60.CisText iB1_RDay 
               Height          =   375
               Left            =   2220
               TabIndex        =   7
               Top             =   60
               Width           =   510
               _ExtentX        =   900
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cCaps           =   2
               cGFormat        =   "###"
               cILength        =   3
            End
            Begin CisText_V60.CisText iB1_TairyuDay 
               Height          =   375
               Left            =   5520
               TabIndex        =   8
               Top             =   60
               Width           =   510
               _ExtentX        =   900
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cCaps           =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D3 
            Height          =   720
            Index           =   1
            Left            =   4575
            Top             =   4500
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   1270
            BackColor       =   14737632
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionX      =   50
            Begin Cis3D_v60.CIS3D CIS3D16 
               Height          =   450
               Index           =   0
               Left            =   60
               Top             =   135
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   794
               BackColor       =   14737632
               Caption         =   "発注かんばん読取　　エラーリスト発行"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   14737632
               cBoderColor2    =   14737632
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cBoderStyle     =   2
            End
            Begin VB.OptionButton H1Op_HKErrPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行有"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   41
               Top             =   135
               Value           =   -1  'True
               Width           =   1050
            End
            Begin VB.OptionButton H1Op_HKErrPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行なし"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   42
               Top             =   375
               Width           =   1170
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D4 
            Height          =   780
            Index           =   1
            Left            =   4575
            Top             =   5235
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   1376
            BackColor       =   14737632
            Caption         =   "発注打切リスト発行"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_HKUchiPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行有"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   43
               Top             =   105
               Value           =   -1  'True
               Width           =   1080
            End
            Begin VB.OptionButton H1Op_HKUchiPrt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発行なし"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   44
               Top             =   375
               Width           =   1080
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   480
            Index           =   11
            Left            =   30
            Top             =   1815
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   847
            BackColor       =   14737632
            Caption         =   "稼働日カレンダー"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_Calendar 
               BackColor       =   &H00E0E0E0&
               Caption         =   "取引先"
               Height          =   255
               Index           =   1
               Left            =   3165
               TabIndex        =   14
               Top             =   105
               Width           =   1185
            End
            Begin VB.OptionButton H1Op_Calendar 
               BackColor       =   &H00E0E0E0&
               Caption         =   "自社"
               Height          =   255
               Index           =   0
               Left            =   1980
               TabIndex        =   13
               Top             =   105
               Value           =   -1  'True
               Width           =   795
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   450
            Index           =   12
            Left            =   30
            Top             =   900
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   794
            BackColor       =   14737632
            Caption         =   "未流動かんばん"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.OptionButton H1Op_Ryudo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "有効"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   9
               Top             =   105
               Value           =   -1  'True
               Width           =   795
            End
            Begin VB.OptionButton H1Op_Ryudo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "無効"
               Height          =   255
               Index           =   1
               Left            =   3510
               TabIndex        =   10
               Top             =   105
               Width           =   765
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   570
            Index           =   9
            Left            =   0
            Top             =   6585
            Width           =   7965
            _ExtentX        =   14049
            _ExtentY        =   1005
            BackColor       =   8421504
            ForeColor       =   16777215
            Caption         =   "《構成展開処理　数量計算》"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cPositionX      =   120
            Begin Cis3D_v60.CIS3D Zai_Pnl 
               Height          =   495
               Index           =   1
               Left            =   3735
               Top             =   45
               Width           =   4185
               _ExtentX        =   7382
               _ExtentY        =   873
               BackColor       =   14737632
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   50
               Begin VB.OptionButton H1Op_Tenkai 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "丸め考慮（収容数/ロット)"
                  Height          =   225
                  Index           =   0
                  Left            =   90
                  TabIndex        =   49
                  Top             =   165
                  Value           =   -1  'True
                  Width           =   2655
               End
               Begin VB.OptionButton H1Op_Tenkai 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "正味数量"
                  Height          =   225
                  Index           =   1
                  Left            =   2955
                  TabIndex        =   50
                  Top             =   150
                  Width           =   1200
               End
            End
         End
      End
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   405
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   714
         BackColor       =   16711680
         ForeColor       =   16777215
         Caption         =   "【 システム共通設定 】"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
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
   End
   Begin Cis3D_v60.CIS3D CIS3D2 
      Height          =   675
      Left            =   -15
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1191
      ForeColor       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HGSｺﾞｼｯｸE"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   1
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   405
         Left            =   3270
         Top             =   225
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   714
         ForeColor       =   0
         Caption         =   "≪ オプション設定 ≫"
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HGSｺﾞｼｯｸE"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   12632256
         cBoderColor2    =   12632256
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderStyle     =   2
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "SIMPR-A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   825
         Index           =   1
         Left            =   240
         TabIndex        =   111
         Top             =   -60
         Width           =   2730
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "SIMPR-A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   825
         Index           =   1
         Left            =   360
         TabIndex        =   112
         Top             =   -30
         Width           =   2730
      End
   End
End
Attribute VB_Name = "CSYSOPTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  システムオプション設定
'**       フォームID    :  COPTION
'**
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/07  By CIS
'**       変  更  日    :  2004/11/12  BY Cis 受注関連（売上処理)追加
'**       変  更  日    :  2004/11/15  By CIS 売上元データ出荷Or検収に対応
'**       変  更  日    :  2005/02/21  By CIS チケット環境（ファイル出力先・機番）追加
'**                                           納入日・便算出方法追加
'**       変  更  日    :  2005/06/18  By CIS かんばん流動履歴日数　追加
'**       変  更  日    :  2005/11/04  By CIS 受注かんばん読取エラーリスト発行有無、Topps空出荷指示書発行　追加
'**                                             発注かんばん読取エラーリスト発行有無・発注打切リスト発行有無　追加
'**       変  更  日    :  2005/11/15  By CIS 稼働日カレンダーの追加
'**       変  更  日    :  2006/02/02  By CIS 未流動かんばんの追加
'**                                           INIファイル-クライアント対応
'**       変  更  日    :  2006/02/22  By CIS 自給チケット出力先追加(ｸﾗｲｱﾝﾄ単位)
'**       変  更  日    :  2006/02/24  By CIS INIファイル-クライアント対応（入力対応がしていなかった)
'**       変  更  日    :  2006/03/16  By CIS ｵﾌﾟｼｮﾝ設定によりﾒﾆｭｰの表示･非表示を行う
'**       変  更  日    :  2006/03/28  By CIS かんばん滞留期間項目追加
'**       変  更  日    :  2006/12/19  By CIS 受注画面入力初期値項目追加
'**       変  更  日    :  2010/06/28  By CIS 展開数量 項目追加
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    Dim Rg_On           As String
    Dim Rg_Off          As String
    Dim wFlg            As Boolean
    Dim wI              As Byte

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End

    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me

    Timer1.Interval = 1000          ' 時刻表示間隔
    Timer1.Enabled = True           ' 時刻表示開始
    B1_Area1.BackColor = gPGotSel
    B1_Area2.BackColor = gPGotSel
'    Me.Move 0, 0, 15300, 9660              '2010/06/28
    Me.Move 0, 0, 15300, 10300              '2010/06/28

'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not gRtn Then
        If Not IniGet Then End
    End If

'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Call CisFrm.InitFld

'   #------------------#
'   #   画面中央表示    #
'   #------------------#
    Call CisFrm.Frm_Center(False)
    
    mGotColor = &HC0FFC0
    mLostColor = &HE0E0E0
'   ( 内容表示 )
    Call BodyDisp
    Call H1chk_Reg_Click

'   ( 組込起動 )
    If Not gRtn Then Me.Show Else PB_END.BtnCaption = "戻る"
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Insert Start
'   受注画面入力 初期値エリア移動
    CIS3D19.Top = CIS3D14(1).Top
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Insert End
'   ( タイトルバー用 )
    With CisFrm
        .TimeSet_TitleName = "SIMPR-A システムオプション設定"
        .TimeSet_TitleSpace = 70
    End With
    
    FormAct = True
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down") Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Function Key_Acc(Key_Code As Integer, IVENT As String)
    If Not FormAct Then Exit Function
    
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    
    Select Case Key_Code
        Case vbKeyF12:      If PB_Ent.Visible Then Call PB_END_Click    '【終了】
        Case vbKeyReturn:   If PB_Ent.Visible Then Call PB_ENT_Click    '【入力】
        Case Else:          Exit Function
    End Select
    
Key_Acc_ED:
    Key_Acc = True
End Function

'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+---------------------+
'+    終了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If Not gRtn Then
        Call CisDB.DBDISConnect
        Call CisFrm.UnLoadDisp
        End
    Else
        Me.Hide
    End If
End Sub

'+-------------------------------+
'+   かんばん印刷先設定
'+-------------------------------+
Private Sub PB_PrtSet_Click(Index As Integer)
    If Index = 0 Then
        Psel_NowSelect = lb_KPrinter
    Else
        Psel_NowSelect = lb_CPrinter
    End If
    
    frmPrintSelect.Show vbModal
    Unload frmPrintSelect
    Set frmPrintSelect = Nothing
    
    If Psel_ChgFlg Then
        If Index = 0 Then
            lb_KPrinter = Psel_ChgPrtName
        Else
            lb_CPrinter = Psel_ChgPrtName
        End If
    End If
End Sub

'+-------------------------------+
'+   通信ポート設定
'+-------------------------------+
Private Sub PB_Set_Click(Index As Integer)
    IRN_LinkNo = Index
    IRN_LinkStr = ConnectionStr(Index)
    
    CSYSPort.Show vbModal
    Unload CSYSPort
    Set CSYSPort = Nothing
End Sub

'+-------------------------------+
'+   自社伝票ファイル出力先　参照　+
'+-------------------------------+
Private Sub PB_JDenFDir_Click()
    With CisFun
        .Folder_Title = "自社伝票ファイル出力先参照"
        .Folder_DefDir = iB1_JDenFDir
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_JDenFDir = RTrim$(.Folder_Result)
    End With
    iB1_JDenFDir.SetFocus
End Sub
'+-------------------------------+
'+   JAMA伝票ファイル出力先　参照　+
'+-------------------------------+
Private Sub PB_JamaFDir_Click()
    With CisFun
        .Folder_Title = "JAMA伝票ファイル出力先参照"
        .Folder_DefDir = iB1_JamaFDir
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_JamaFDir = RTrim$(.Folder_Result)
    End With
    iB1_JamaFDir.SetFocus
End Sub
'+-----------------------------------+
'+   自給チケットファイル出力先　参照　+
'+-----------------------------------+
Private Sub PB_JTicketFDir_Click()
    With CisFun
        .Folder_Title = "自給チケットファイル出力先参照"
        .Folder_DefDir = iB1_JTicketFDir
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_JTicketFDir = RTrim$(.Folder_Result)
    End With
    iB1_JTicketFDir.SetFocus
End Sub
'+-----------------------------------+
'+   支給チケットファイル出力先　参照　+
'+-----------------------------------+
Private Sub PB_STicketFDir_Click()
    With CisFun
        .Folder_Title = "支給チケットファイル出力先参照"
        .Folder_DefDir = iB1_STicketFDir
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_STicketFDir = RTrim$(.Folder_Result)
    End With
    iB1_STicketFDir.SetFocus
End Sub
'+-------------------------------+
'+   平準化月超過ファイル出力先　参照　+
'+-------------------------------+
Private Sub PB_HeiFDir_Click()
    With CisFun
        .Folder_Title = "平準化月超過ファイル出力先参照"
        .Folder_DefDir = iB1_HeiFDir
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_HeiFDir = RTrim$(.Folder_Result)
    End With
    iB1_HeiFDir.SetFocus
End Sub
'+-------------------------------+
'+   内示変換TEXT格納先　参照　+
'+-------------------------------+
Private Sub PB_TextDir_Click()
    With CisFun
'        .Folder_Title = "内示変換TEXT格納先参照"   ' 2004/11/08 MNT
        .Folder_Title = "TEXT変換元格納先参照"
        .Folder_DefDir = iB1_TextDir
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_TextDir = RTrim$(.Folder_Result)
    End With
    iB1_TextDir.SetFocus
End Sub
'+-----------------------------------+
'+  自社ｺｰﾄﾞﾁｪｯｸ
'+-----------------------------------+
Private Sub H1Ck_Jisya_Click()
    If H1Ck_Jisya.Value = 0 Then
       H1Ck_Jisya.Caption = "無"
    Else
       H1Ck_Jisya.Caption = "有"
    End If
End Sub
'+-------------------------------+
'+   イメージ格納先　参照　+
'+-------------------------------+
Private Sub PB_ImageDir_Click()
    With CisFun
        .Folder_Title = "イメージ格納先参照"
        .Folder_DefDir = iB1_ImageDir
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_ImageDir = RTrim$(.Folder_Result)
    End With
    iB1_ImageDir.SetFocus
End Sub
'////////////////////////////
'//     受注関連
'////////////////////////////

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Insert Start
'+-----------------------------------+
'+  受注画面入力初期値(H1Op_JGamen)
'+-----------------------------------+
Private Sub H1Op_JGamen_Click(Index As Integer)
    H1Op_JGamen(0).Tag = Format(Index, "0")
    H1Op_JGamen(0).BackColor = mLostColor
    H1Op_JGamen(1).BackColor = mLostColor
    H1Op_JGamen(Index).BackColor = mGotColor
End Sub
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Insert End

'+-----------------------------------+
'+  売上処理(H1Op_Uriage)
'+-----------------------------------+
Private Sub H1Op_Uriage_Click(Index As Integer)
    H1Op_Uriage(0).Tag = Format(Index, "0")
    H1Op_Uriage(0).BackColor = mLostColor
    H1Op_Uriage(1).BackColor = mLostColor
    H1Op_Uriage(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  受注かんばん読取りｴﾗｰﾘｽﾄ(H1Op_JKErrPrt)
'+-----------------------------------+
Private Sub H1Op_JKErrPrt_Click(Index As Integer)
    H1Op_JKErrPrt(0).Tag = Format(Index, "0")
    H1Op_JKErrPrt(0).BackColor = mLostColor
    H1Op_JKErrPrt(1).BackColor = mLostColor
    H1Op_JKErrPrt(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  Topps空出荷指示書(H1Op_JToppsPrt)
'+-----------------------------------+
Private Sub H1Op_JToppsPrt_Click(Index As Integer)
    H1Op_JToppsPrt(0).Tag = Format(Index, "0")
    H1Op_JToppsPrt(0).BackColor = mLostColor
    H1Op_JToppsPrt(1).BackColor = mLostColor
    H1Op_JToppsPrt(Index).BackColor = mGotColor
End Sub
'////////////////////////////
'//     発注関連
'////////////////////////////
'+-----------------------------------+
'+  臨時かんばん発行(H1Op_RinPrt)
'+-----------------------------------+
Private Sub H1Op_RinPrt_Click(Index As Integer)
    H1Op_RinPrt(0).Tag = Format(Index, "0")
    H1Op_RinPrt(0).BackColor = mLostColor
    H1Op_RinPrt(1).BackColor = mLostColor
    H1Op_RinPrt(Index).BackColor = mGotColor
End Sub
'------------------------------------------------------- 2006/02/02 Insert
'+---------------------------------------------+
'+  未流動かんばん　(H1Op_Ryudo)
'+---------------------------------------------+
Private Sub H1Op_Ryudo_Click(Index As Integer)
    H1Op_Ryudo(0).Tag = Format(Index, "0")
    H1Op_Ryudo(0).BackColor = mLostColor
    H1Op_Ryudo(1).BackColor = mLostColor
    H1Op_Ryudo(Index).BackColor = mGotColor
End Sub
'------------------------------------------------------- 2006/02/02 Insert End
'+-----------------------------------+
'+  現品票発行(H1Op_GenPrt)
'+-----------------------------------+
Private Sub H1Op_GenPrt_Click(Index As Integer)
    H1Op_GenPrt(0).Tag = Format(Index, "0")
    H1Op_GenPrt(0).BackColor = mLostColor
    H1Op_GenPrt(1).BackColor = mLostColor
    H1Op_GenPrt(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  端数ラベル発行(H1Op_HasPrt)
'+-----------------------------------+
Private Sub H1Op_HasPrt_Click(Index As Integer)
    H1Op_HasPrt(0).Tag = Format(Index, "0")
    H1Op_HasPrt(0).BackColor = mLostColor
    H1Op_HasPrt(1).BackColor = mLostColor
    H1Op_HasPrt(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  発注明細（控）発行(H1Op_HIkPrt)
'+-----------------------------------+
Private Sub H1Op_HIkPrt_Click(Index As Integer)
    H1Op_HIkPrt(0).Tag = Format(Index, "0")
    H1Op_HIkPrt(0).BackColor = mLostColor
    H1Op_HIkPrt(1).BackColor = mLostColor
    H1Op_HIkPrt(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  発注明細（空明細）発行(H1Op_KarPrt)
'+-----------------------------------+
Private Sub H1Op_KarPrt_Click(Index As Integer)
    H1Op_KarPrt(0).Tag = Format(Index, "0")
    H1Op_KarPrt(0).BackColor = mLostColor
    H1Op_KarPrt(1).BackColor = mLostColor
    H1Op_KarPrt(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  発注かんばん読取りｴﾗｰﾘｽﾄ(H1Op_HKErrPrt)
'+-----------------------------------+
Private Sub H1Op_HKErrPrt_Click(Index As Integer)
    H1Op_HKErrPrt(0).Tag = Format(Index, "0")
    H1Op_HKErrPrt(0).BackColor = mLostColor
    H1Op_HKErrPrt(1).BackColor = mLostColor
    H1Op_HKErrPrt(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  発注打切りﾘｽﾄ(H1Op_HKUchiPrt)
'+-----------------------------------+
Private Sub H1Op_HKUchiPrt_Click(Index As Integer)
    H1Op_HKUchiPrt(0).Tag = Format(Index, "0")
    H1Op_HKUchiPrt(0).BackColor = mLostColor
    H1Op_HKUchiPrt(1).BackColor = mLostColor
    H1Op_HKUchiPrt(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  受領印
'+-----------------------------------+
Private Sub H1Ck_JHanko_Click()
    If H1Ck_JHanko.Value = 0 Then
       H1Ck_JHanko.Caption = "印字なし"
       H1Ck_JDate.Visible = False
       JUe.Visible = False
       JShita.Visible = False
       iB1_JUe.Visible = False
       iB1_JShita.Visible = False
    Else
       H1Ck_JHanko.Caption = "印字あり"
       H1Ck_JDate.Visible = True
       JUe.Visible = True
       JShita.Visible = True
       iB1_JUe.Visible = True
       iB1_JShita.Visible = True
    End If
End Sub
'+-----------------------------------+
'+  受領印
'+-----------------------------------+
Private Sub H1Ck_JDate_Click()
    If H1Ck_JDate.Value = 0 Then
       H1Ck_JDate.Caption = "受領日印字あり"
    Else
       H1Ck_JDate.Caption = "受領日印字なし"
    End If
End Sub
'+-----------------------------------+
'+  かんばん読取り 即時発注(H1Op_KPool)
'+-----------------------------------+
Private Sub H1Op_KPool_Click(Index As Integer)
    H1Op_KPool(0).Tag = Format(Index, "0")
    H1Op_KPool(0).BackColor = mLostColor
    H1Op_KPool(1).BackColor = mLostColor
    H1Op_KPool(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  平準化月超過発注(H1Op_KHei)
'+-----------------------------------+
Private Sub H1Op_KHei_Click(Index As Integer)
    H1Op_KHei(0).Tag = Format(Index, "0")
    H1Op_KHei(0).BackColor = mLostColor
    H1Op_KHei(1).BackColor = mLostColor
    H1Op_KHei(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  平準化月超過発注出力先
'+-----------------------------------+
Private Sub H1Ck_KHei_Click()
    If H1Ck_KHei.Value = 0 Then
       H1Ck_KHei.Caption = "ＣＳＶ"
    Else
       H1Ck_KHei.Caption = "帳票"
    End If
End Sub
'+-----------------------------------+
'+  材料発注管理システム    登録
'+-----------------------------------+
Private Sub H1chk_Reg_Click()
    If H1chk_Reg.Value = 1 Then
       H1chk_Reg.Caption = Rg_On
    Else
       H1chk_Reg.Caption = Rg_Off
    End If

    For wI = 0 To 0
        Zai_Pnl(wI).Enabled = (H1chk_Reg = 1)
    Next wI
End Sub
'+---------------------------------------------+
'+  材料発注　内示対象データ　(H1Op_ZNJYMKB)
'+---------------------------------------------+
Private Sub H1Op_ZNJYMKB_Click(Index As Integer)
    H1Op_ZNJYMKB(0).Tag = Format(Index, "0")
    H1Op_ZNJYMKB(0).BackColor = mLostColor
    H1Op_ZNJYMKB(1).BackColor = mLostColor
    H1Op_ZNJYMKB(2).BackColor = mLostColor
    H1Op_ZNJYMKB(Index).BackColor = mGotColor
End Sub
'=============================== 2010/06/28 Start
'+---------------------------------------------+
'+  構成展開　処理数量　(H1Op_Tenkai)
'+---------------------------------------------+
Private Sub H1Op_Tenkai_Click(Index As Integer)
    H1Op_Tenkai(0).Tag = Format(Index, "0")
    H1Op_Tenkai(0).BackColor = mLostColor
    H1Op_Tenkai(1).BackColor = mLostColor
    H1Op_Tenkai(Index).BackColor = mGotColor
End Sub
'=============================== 2010/06/28 End
'+---------------------------------------------+
'+  稼働日カレンダー　(H1Op_Calendar)
'+---------------------------------------------+
Private Sub H1Op_Calendar_Click(Index As Integer)
    H1Op_Calendar(0).Tag = Format(Index, "0")
    H1Op_Calendar(0).BackColor = mLostColor
    H1Op_Calendar(1).BackColor = mLostColor
    H1Op_Calendar(Index).BackColor = mGotColor
End Sub
'+---------------------------------------------+
'+  納入日・便算出方法　(H1Op_NYmdBin)
'+---------------------------------------------+
Private Sub H1Op_NYmdBin_Click(Index As Integer)
    H1Op_NYmdBin(0).Tag = Format(Index, "0")
    H1Op_NYmdBin(0).BackColor = mLostColor
    H1Op_NYmdBin(1).BackColor = mLostColor
    H1Op_NYmdBin(Index).BackColor = mGotColor
End Sub
'+---------------------------------------------+
'+  完納予定日・便算出方法　(H1Op_KYmd)
'+---------------------------------------------+
Private Sub H1Op_KYmd_Click(Index As Integer)
    H1Op_KYmd(0).Tag = Format(Index, "0")
    H1Op_KYmd(0).BackColor = mLostColor
    H1Op_KYmd(1).BackColor = mLostColor
    H1Op_KYmd(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  かんばんｿｰﾀｰ(H1Op_KiKi1)
'+-----------------------------------+
Private Sub H1Op_KiKi1_Click(Index As Integer)
    H1Op_KiKi1(0).Tag = Format(Index, "0")
    H1Op_KiKi1(0).BackColor = mLostColor
    H1Op_KiKi1(1).BackColor = mLostColor
    H1Op_KiKi1(Index).BackColor = mGotColor
    If Index = 1 Then
        PB_Set(0).Visible = False
    Else
        PB_Set(0).Visible = True
    End If
End Sub
'+-----------------------------------+
'+  かんばん手差ﾘｰﾀﾞｰ(H1Op_KiKi2)
'+-----------------------------------+
Private Sub H1Op_KiKi2_Click(Index As Integer)
    H1Op_KiKi2(0).Tag = Format(Index, "0")
    H1Op_KiKi2(0).BackColor = mLostColor
    H1Op_KiKi2(1).BackColor = mLostColor
    H1Op_KiKi2(Index).BackColor = mGotColor
    If Index = 1 Then
        PB_Set(1).Visible = False
    Else
        PB_Set(1).Visible = True
    End If
End Sub
'+-----------------------------------+
'+  かんばんﾌﾟﾘﾝﾀ(H1Op_KiKi3)
'+-----------------------------------+
Private Sub H1Op_KiKi3_Click(Index As Integer)
    H1Op_KiKi3(0).Tag = Format(Index, "0")
    H1Op_KiKi3(0).BackColor = mLostColor
    H1Op_KiKi3(1).BackColor = mLostColor
    H1Op_KiKi3(Index).BackColor = mGotColor
    If Index = 1 Then
        PB_Set(2).Visible = False
    Else
        PB_Set(2).Visible = True
    End If
End Sub
'+-----------------------------------+
'+  QRﾊﾝﾄﾞｽｷｬﾅ(H1Op_KiKi4)
'+-----------------------------------+
Private Sub H1Op_KiKi4_Click(Index As Integer)
    H1Op_KiKi4(0).Tag = Format(Index, "0")
    H1Op_KiKi4(0).BackColor = mLostColor
    H1Op_KiKi4(1).BackColor = mLostColor
    H1Op_KiKi4(Index).BackColor = mGotColor
    If Index = 1 Then
        PB_Set(3).Visible = False
    Else
        PB_Set(3).Visible = True
    End If
End Sub
'+-----------------------------------+
'+  ﾊﾞｰｺｰﾄﾞ読取装置(ｷｰﾎﾞｰﾄﾞｲﾝﾀｰﾌｪｲｽ）(H1Op_KiKi5)
'+-----------------------------------+
Private Sub H1Op_KiKi5_Click(Index As Integer)
    H1Op_KiKi5(0).Tag = Format(Index, "0")
    H1Op_KiKi5(0).BackColor = mLostColor
    H1Op_KiKi5(1).BackColor = mLostColor
    H1Op_KiKi5(Index).BackColor = mGotColor
End Sub
'+-----------------------------------+
'+  かんばん手差ﾘｰﾀﾞｰ(H1Op_KiKi6)
'+-----------------------------------+
Private Sub H1Op_KiKi6_Click(Index As Integer)
    H1Op_KiKi6(0).Tag = Format(Index, "0")
    H1Op_KiKi6(0).BackColor = mLostColor
    H1Op_KiKi6(1).BackColor = mLostColor
    H1Op_KiKi6(Index).BackColor = mGotColor
    If Index = 1 Then
        PB_Set(5).Visible = False
    Else
        PB_Set(5).Visible = True
    End If
End Sub
'+-----------------------------------+
'+  QRﾎｯﾊﾟﾘｰﾀﾞｰ(H1Op_KiKi7)
'+-----------------------------------+
Private Sub H1Op_KiKi7_Click(Index As Integer)
    H1Op_KiKi7(0).Tag = Format(Index, "0")
    H1Op_KiKi7(0).BackColor = mLostColor
    H1Op_KiKi7(1).BackColor = mLostColor
    H1Op_KiKi7(Index).BackColor = mGotColor
    If Index = 1 Then
        PB_Set(6).Visible = False
    Else
        PB_Set(6).Visible = True
    End If
End Sub
'+-----------------------------------+
'+  仕切かんばん発行
'+-----------------------------------+
Private Sub H1Ck_SPrint_Click()
    If H1Ck_SPrint.Value = 0 Then
       H1Ck_SPrint.Caption = "なし"
    Else
       H1Ck_SPrint.Caption = "あり"
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If Not BodyChk Then GoTo ReturnPress_Ed
    If Not DBPut Then GoTo ReturnPress_Ed
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      BODYﾁｪｯｸ処理
'****************************
Private Function BodyChk() As Boolean
    Dim wSaveAcc        As String
    BodyChk = False
' ﾕｰｻﾞｰｺｰﾄﾞ入力ﾁｪｯｸ
    If Trim(iB1_HMoto) = "" Then
       MsgBox "発行元を入力して下さい。", vbExclamation, "必須入力ｴﾗｰ"
       iB1_HMoto.SetFocus
       Exit Function
    End If

    BodyChk = True
End Function

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet_TitleSpace = 360
    Call CisFrm.TimeSet
End Sub
'+----------------------+
'+     内 容 表 示      +
'+----------------------+
Private Sub BodyDisp()
'受注関連設定
'<< IRNAKSIC >>

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Insert Start
'受注画面入力 初期値
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 48
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 48
        IRN.oyian = "0"
        IRN.oyian2 = "受注画面入力初期値"
        Call IRNInsert
    End If
    H1Op_JGamen(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Insert End

' 売上処理
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 50
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 50
        IRN.oyian = "0"
        IRN.oyian2 = "売上処理"
        Call IRNInsert
    End If
    H1Op_Uriage(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
' 受注かんばん読取りエラーリスト発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 51
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 51
        IRN.oyian = "0"
        IRN.oyian2 = "受注読取エラーリスト"
        Call IRNInsert
    End If
    H1Op_JKErrPrt(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
' Topps空出荷指示書発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 52
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 52
        IRN.oyian = "0"
        IRN.oyian2 = "Topps空出荷指示書"
        Call IRNInsert
    End If
    H1Op_JToppsPrt(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
'発注関連設定
'<< IRNAKSIC >>
' 臨時かんばん発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 1
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 1
        IRN.oyian = "0"
        IRN.oyian2 = "臨時かんばん"
        Call IRNInsert
    End If
    H1Op_RinPrt(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
' 現品票発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 2
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 2
        IRN.oyian = "0"
        IRN.oyian2 = "現品票"
        Call IRNInsert
    End If
    H1Op_GenPrt(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
' 端数ラベル発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 3
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 3
        IRN.oyian = "0"
        IRN.oyian2 = "端数ラベル"
        Call IRNInsert
    End If
    H1Op_HasPrt(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
' 発注明細控
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 12
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 12
        IRN.oyian = "0"
        IRN.oyian2 = "発注明細【控】"
        Call IRNInsert
    End If
    H1Op_HIkPrt(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
' 発注明細空明細
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 13
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 13
        IRN.oyian = "0"
        IRN.oyian2 = "発注明細【空明細】"
        Call IRNInsert
    End If
    H1Op_KarPrt(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
    
' 発注かんばん読取りエラーリスト発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 35
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 35
        IRN.oyian = "0"
        IRN.oyian2 = "発注読取エラーリスト"
        Call IRNInsert
    End If
    H1Op_HKErrPrt(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
' 発注打切りリスト発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 36
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 36
        IRN.oyian = "0"
        IRN.oyian2 = "発注打切りリスト"
        Call IRNInsert
    End If
    H1Op_HKUchiPrt(CisFun.Val2(Left(IRN.oyian, 1))).Value = True

' 受領印
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 20
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 20
        IRN.oyian = "00受領"
        IRN.oyian2 = ""
        Call IRNInsert
    End If
    H1Ck_JHanko.Value = CisFun.Val2(Left(IRN.oyian, 1))
    H1Ck_JDate.Value = CisFun.Val2(Mid(IRN.oyian, 2, 1))
    iB1_JUe = CisFun.Mid2(IRN.oyian, 3, 8)
    iB1_JShita = CisFun.Mid2(IRN.oyian2, 1, 8)
    Call H1Ck_JHanko_Click
    Call H1Ck_JDate_Click
' 発行元
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 4
        IRN.oyian = "SIMPR-A"
        IRN.oyian2 = "発行元"
        Call IRNInsert
    End If
    iB1_HMoto = Mid(IRN.oyian, 1, 7)
' かんばん読取り 即時発注
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 5
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 5
        IRN.oyian = "0"
        IRN.oyian2 = "かんばん即時発注"
        Call IRNInsert
    End If
    H1Op_KPool(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
' 平準化月超過発注
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 6
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 6
        IRN.oyian = "00"
        IRN.oyian2 = "平準化月超過発注"
        Call IRNInsert
    End If
    H1Op_KHei(CisFun.Val2(Mid(IRN.oyian, 1, 1))).Value = True
    H1Ck_KHei.Value = CisFun.Val2(Mid(IRN.oyian, 2, 1))
'+---------------------------------------------------------------------  2005/06/18 Insert
' かんばん流動履歴日数
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 32
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 32
        IRN.oyian = "020"
        IRN.oyian2 = "かんばん流動履歴日数"
        Call IRNInsert
    End If
    iB1_RDay = CisFun.RSetFld(Mid(IRN.oyian, 1, 3), 3, "##0")
'+---------------------------------------------------------------------  2005/06/18 Insert End
'+---------------------------------------------------------------------  2006/03/28 Insert
' かんばん滞留期間日数
    ' 滞留日付設定値・重複削除日付設定値を求める
    gSL_Select = " SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'データ管理'"
    gSL_Select = gSL_Select & "   AND 区分タイプ = '' "
    gSL_Select = gSL_Select & "   AND 値 = 'KAN' "
    If Not SYMRead(gSL_Select, 1) Then
        Call ItemsClearSYMCreate
        With SYMCreate
            .区分名称 = "データ管理"
            .区分桁数 = 5
            .値 = "KAN"
            .値名称 = "発注かんばん処理削除管理"
            .英数字1 = "滞留期間(稼動日) (1)"
            .英数字2 = "発注かんばん重複削除日数 (2)"
            .英数字3 = "かんばんﾁｪｯｸﾃｰﾌﾞﾙ削除日数(3)"
            .数字1 = 0
            .数字2 = 1
            .数字3 = 30
            .作成者 = gTanto
            .制御区分 = "X"
        End With
        If SYMInitCreate(True) Then
           Call SYMInitCreate(False)
        End If
        SYM.数字1 = 0
        SYM.数字2 = 1
        SYM.数字3 = 30
    End If
    iB1_TairyuDay = CisFun.RSetFld(SYM.数字1, 3, "##0")
    iB1_TairyuDay.Tag = SYM.数字1
'+---------------------------------------------------------------------  2006/03/28 Insert End
'+---------------------------------------------------------------------  2006/02/02 Insert
' 未流動かんばん
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 37
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 37
        IRN.oyian = "0"
        IRN.oyian2 = "未流動かんばん"
        Call IRNInsert
    End If
    H1Op_Ryudo(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
'+---------------------------------------------------------------------  2006/02/02 Insert End
' 自社伝票発行用(工場・機番)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 7
        IRN.oyian = "0"
        IRN.oyian2 = "00"
        Call IRNInsert
    End If
    ' 工場
    iB1_Kojyo = RTrim(IRN.oyian)
    ' 機番
    iB1_Kiban = RTrim(IRN.oyian2)
' JAMA伝票発行用(工場・機番)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 8
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 8
        IRN.oyian = "0"
        IRN.oyian2 = "00"
        Call IRNInsert
    End If
    ' 工場
    iB1_JKojyo = RTrim(IRN.oyian)
    ' 機番
    iB1_JKiban = RTrim(IRN.oyian2)

' 材料発注管理オプション
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 10
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 10
        IRN.oyian = "0"
        IRN.oyian2 = "材料発注管理オプション"
        Call IRNInsert
    End If

    If IRN.oyian = "1" Then
        Rg_On = "登録済":   Rg_Off = "登録解除"
    Else
        Rg_On = "登録":   Rg_Off = "未登録"
    End If
    H1chk_Reg.Value = CisFun.Val2(IRN.oyian)

' 材料発注　内示対象データ
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 11
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 11
        IRN.oyian = "0"
        IRN.oyian2 = "材料発注　内示対象データ"
        Call IRNInsert
    End If
    H1Op_ZNJYMKB(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
'============================== 2010/06/28 Start
' 構成展開 計算数量
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 53
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 53
        IRN.oyian = "0"
        IRN.oyian2 = "構成展開 計算数量"
        Call IRNInsert
    End If
    H1Op_Tenkai(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
'============================== 2010/06/28 End
'----------------------------------------------------- 2005.02.21 追加 ( Start )----------'
' 納入日・便算出方法
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 14
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 14
        IRN.oyian = "0"
        Call IRNInsert
    End If
    H1Op_NYmdBin(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
    H1Op_NYmdBin(0).Tag = Left(IRN.oyian, 1)
' チケット伝票発行用(機番)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 15
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 15
        IRN.oyian = ""
        IRN.oyian2 = "00"
        Call IRNInsert
    End If
    ' 機番
    iB1_CKiban = RTrim(IRN.oyian2)
'----------------------------------------------------- 2005.02.21 追加 ( End   )----------'
'----------------------------------------------------- 2005.03.04 追加 ( Start )----------'
' 観桜予定日・便算出方法
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 16
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 16
        IRN.oyian = "0"
        Call IRNInsert
    End If
    H1Op_KYmd(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
    H1Op_KYmd(0).Tag = Left(IRN.oyian, 1)
'----------------------------------------------------- 2005.03.04 追加 ( End )----------'
' かんばん設定　仕切かんばん
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 31
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 31
        IRN.oyian = "0"
        IRN.oyian2 = "仕切かんばん"
        Call IRNInsert
    End If
    H1Ck_SPrint.Value = CisFun.Val2(Mid(IRN.oyian, 1, 1))
    iB1_SHeigh = CisFun.RSetFld(CisFun.Val2(Mid(IRN.oyian, 2, 2)), 2, "##")
'----------------------------------------------------- 2005.11.15 追加 ( Start )----------'
' 稼働日カレンダー
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 33
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 33
        IRN.oyian = "0"
        IRN.oyian2 = "稼働日カレンダー"
        Call IRNInsert
    End If
    H1Op_Calendar(CisFun.Val2(Left(IRN.oyian, 1))).Value = True
    H1Op_Calendar(0).Tag = Left(IRN.oyian, 1)
'----------------------------------------------------- 2005.11.15 追加 ( End )----------'
    
'<< INI >>
    With CisFun
'                                           < 2006/02/24 MNT START >
'        .INI_FileID = gIniExe & gIniName
        .INI_FileID = gIniExe & gwsIniName
'                                           < 2006/02/24 MNT  END  >
        .INI_Section = "Option"
    ' 自社伝票ファイル出力先
        .INI_Key = "DENPYOFILEOUT":       .INI_Default = gIniExe
        .GetIni
        iB1_JDenFDir = RTrim(.INI_String)
    ' JAMA伝票ファイル出力先
        .INI_Key = "JAMAFILEOUT":       .INI_Default = gIniExe
        .GetIni
        iB1_JamaFDir = RTrim(.INI_String)
'----------------------------------------------------- 2005.02.21 追加 ( Start )----------'
    ' 自給チケットファイル出力先
        .INI_Key = "JTICKETFILEOUT":       .INI_Default = gIniExe
        .GetIni
        iB1_JTicketFDir = RTrim(.INI_String)
    ' 支給チケットファイル出力先
        .INI_Key = "STICKETFILEOUT":       .INI_Default = gIniExe
        .GetIni
        iB1_STicketFDir = RTrim(.INI_String)
'----------------------------------------------------- 2005.02.21 追加 ( End   )----------'
    ' 平準化月超過ファイル出力先
        .INI_Key = "HEIFILEOUT":       .INI_Default = gIniExe
        .GetIni
        iB1_HeiFDir = RTrim(.INI_String)
    ' イメージ格納先
        .INI_Key = "Image":       .INI_Default = gIniExe
        .GetIni
        iB1_ImageDir = RTrim(.INI_String)
    ' 接続機器
        For gInt = 0 To 6
'                                               < 2006/02/02 ADD START >
            .INI_FileID = gIniExe & gIniName
'                                               < 2006/02/02 ADD  END  >
            .INI_Section = "Apparatus"
            .INI_Key = gInt + 1: .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               .INI_Key = gInt + 1
               .INI_String = Connection(gInt)
               .PutIni
'                                               < 2006/02/02 ADD START >
                .INI_FileID = gIniExe & gwsIniName
'                                               < 2006/02/02 ADD  END  >
               .INI_Section = Connection(gInt)
               .INI_Key = "Use"
               .INI_String = "Yes"
               .PutIni
            End If
'                                               < 2006/02/02 ADD START >
            .INI_FileID = gIniExe & gwsIniName
'                                               < 2006/02/02 ADD  END  >
            .INI_Section = Connection(gInt)
            .INI_Key = "Use": .INI_Default = ""
            .GetIni
            If UCase(Trim(.INI_String)) = "N" Or UCase(Trim(.INI_String)) = "NO" Then
               gCnt1 = "1"
            Else
               gCnt1 = "0"
            End If
            Select Case gInt
                Case 0: H1Op_KiKi1(gCnt1).Value = True
                Case 1: H1Op_KiKi2(gCnt1).Value = True
                Case 2: H1Op_KiKi3(gCnt1).Value = True
                Case 3: H1Op_KiKi4(gCnt1).Value = True
                Case 4: H1Op_KiKi5(gCnt1).Value = True
                Case 5: H1Op_KiKi6(gCnt1).Value = True
                Case 6: H1Op_KiKi7(gCnt1).Value = True
            End Select
        Next gInt
        ' 変換元フォルダ
        .INI_Section = "TextConvert"
        .INI_Default = gIniExe & "Text\"
        .INI_Key = "InputDir"
        .GetIni
        iB1_TextDir = RTrim(.INI_String)
        .INI_Key = "LogDelDay"
        .INI_Default = "60"
        .GetIni
        iB1_LogDel = CisFun.Val2(.INI_String)
        .INI_Key = "Jisya"
        .INI_Default = "No"
        .GetIni
        If UCase(.INI_String) = "Y" Or UCase(.INI_String) = "YES" Then
           H1Ck_Jisya.Value = 1
        Else
           H1Ck_Jisya.Value = 0
        End If
        
        ' かんばん印刷先
        .INI_Section = "かんばん印刷先"
        .INI_Default = ""
        .INI_Key = CisFun.WSID
        .GetIni
        If Trim(.INI_String) <> "" Then
           lb_KPrinter = RTrim(.INI_String)
        End If
        
'-----     -----     -----     -----     -----     -----     < 2006/02/22 ADD START >
'
'自給チケット出力先
'
        .INI_Section = "自給チケット印刷先"
        .INI_Default = ""
        .INI_Key = CisFun.WSID
        .GetIni
        If Trim(.INI_String) <> "" Then lb_CPrinter = RTrim(.INI_String)
'
'-----     -----     -----     -----     -----     -----     < 2006/02/22 ADD  END  >
'                                           < 2006/02/24 MNT START >
        .INI_FileID = gIniExe & gIniName
'                                           < 2006/02/24 MNT  END  >
    End With
End Sub
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()

    DBPut = False
'  ( 確認メッセージ )
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "     【 オプション設定変更確認 】      "
        .MB_MSG(3) = "    オプション設定を変更しますか      "
        .MB_Title = "オプション設定"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Insert Start
' 受注画面入力初期値
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 48
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_JGamen(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_JGamen(0).Tag)
       Call IRNInsert
    End If
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Insert End

' 売上処理
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 50
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_Uriage(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_Uriage(0).Tag)
       Call IRNInsert
    End If
' 受注かんばん読取りエラーリスト
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 51
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_JKErrPrt(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_JKErrPrt(0).Tag)
       Call IRNInsert
    End If
' Topps出荷指示書リスト
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 52
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_JToppsPrt(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_JToppsPrt(0).Tag)
       Call IRNInsert
    End If
' 臨時かんばん発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 1
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_RinPrt(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_RinPrt(0).Tag)
       Call IRNInsert
    End If
' 現品票発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 2
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_GenPrt(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_GenPrt(0).Tag)
       Call IRNInsert
    End If
' 端数ラベル発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 3
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_HasPrt(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_HasPrt(0).Tag)
       Call IRNInsert
    End If
' 発注明細控
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 12
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_HIkPrt(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_HIkPrt(0).Tag)
       Call IRNInsert
    End If
' 発注明細空明細
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 13
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_KarPrt(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_KarPrt(0).Tag)
       Call IRNInsert
    End If

' 発注かんばん読取りエラーリスト
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 35
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_HKErrPrt(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_HKErrPrt(0).Tag)
       Call IRNInsert
    End If
' 発注打切りリスト
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 36
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_HKUchiPrt(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_HKUchiPrt(0).Tag)
       Call IRNInsert
    End If

' 受領印
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 20
    If IRNRead("", 1) Then
       IRN.oyian = Format(H1Ck_JHanko.Value, "0") & _
                   Format(H1Ck_JDate.Value, "0") & _
                   iB1_JUe
       IRN.oyian2 = iB1_JShita
       Call IRNUpdate
    Else
       IRN.oyian = Format(H1Ck_JHanko.Value, "0") & _
                   Format(H1Ck_JDate.Value, "0") & _
                   iB1_JUe
       IRN.oyian2 = iB1_JShita
       Call IRNInsert
    End If
' 発行元
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    If IRNRead("", 1) Then
       IRN.oyian = RTrim(iB1_HMoto)
       Call IRNUpdate
    Else
       IRN.oyian = RTrim(iB1_HMoto)
       Call IRNInsert
    End If
'  かんばん読取り 即時発注
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 5
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_KPool(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_KPool(0).Tag)
       Call IRNInsert
    End If
'  平準化月超過発注
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 6
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_KHei(0).Tag) & H1Ck_KHei.Value
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_KHei(0).Tag) & H1Ck_KHei.Value
       Call IRNInsert
    End If
'+-------------------------------------------------------------------   2005/06/18 Insert
'  平準化月超過発注
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 32
    If IRNRead("", 1) Then
       IRN.oyian = Format(CisFun.Val2(iB1_RDay), "000")
       Call IRNUpdate
    Else
       IRN.oyian = Format(CisFun.Val2(iB1_RDay), "000")
       Call IRNInsert
    End If
'+-------------------------------------------------------------------   2005/06/18 Insert End
'+-------------------------------------------------------------------   2006/03/28 Insert
'  かんばん滞留日数
    If CisFun.Val2(iB1_TairyuDay.Tag) <> CisFun.Val2(iB1_TairyuDay) Then
       With CisDB
            .SQL = "Update 名称マスタ Set "
            .SQL = .SQL & " 数字1 = " & CisFun.Val2(iB1_TairyuDay) & ","
            .SQL = .SQL & "更新日 = GETDATE(),"
            .SQL = .SQL & "更新者 = '" & RTrim(gTanto) & "',"
            .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20)"
            .SQL = .SQL & " WHERE 区分名称 = 'データ管理'"
            .SQL = .SQL & "   AND 区分タイプ = '' "
            .SQL = .SQL & "   AND 値 = 'KAN' "
            Call .DBExec
        End With
        iB1_TairyuDay.Tag = iB1_TairyuDay
    End If
'+-------------------------------------------------------------------   2006/03/28 Insert End
'+-------------------------------------------------------------------   2006/02/02 Insert
'  未流動かんばん
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 37
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_Ryudo(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_Ryudo(0).Tag)
       Call IRNInsert
    End If
'+-------------------------------------------------------------------   2006/02/02 Insert End

'  自社伝票発行用(工場・機番)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    If IRNRead("", 1) Then
       IRN.oyian = Trim(iB1_Kojyo)
       IRN.oyian2 = Trim(iB1_Kiban)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(iB1_Kojyo)
       IRN.oyian2 = Trim(iB1_Kiban)
       Call IRNInsert
    End If
'  JAMA伝票発行用(工場・機番)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 8
    If IRNRead("", 1) Then
       IRN.oyian = Trim(iB1_JKojyo)
       IRN.oyian2 = Trim(iB1_JKiban)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(iB1_JKojyo)
       IRN.oyian2 = Trim(iB1_JKiban)
       Call IRNInsert
    End If
    
' 材料発注管理オプション
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 10
    If IRNRead("", 1) Then
       IRN.oyian = H1chk_Reg.Value
       Call IRNUpdate
    Else
       IRN.oyian = H1chk_Reg.Value
       Call IRNInsert
    End If
    
' 材料発注　内示対象データ
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 11
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_ZNJYMKB(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_ZNJYMKB(0).Tag)
       Call IRNInsert
    End If
'=========================================== 2010/06/28 Start
' 構成展開　計算数量
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 53
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_Tenkai(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_Tenkai(0).Tag)
       Call IRNInsert
    End If
'=========================================== 2010/06/28 End
'----------------------------------------------------- 2005.02.21 追加 ( Start )----------'
' 納入日・便算出方法
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 14
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_NYmdBin(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_NYmdBin(0).Tag)
       Call IRNInsert
    End If
'  チケット発行用(機番)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 15
    If IRNRead("", 1) Then
       IRN.oyian2 = Trim(iB1_CKiban)
       Call IRNUpdate
    Else
       IRN.oyian2 = Trim(iB1_CKiban)
       Call IRNInsert
    End If
'----------------------------------------------------- 2005.02.21 追加 ( End   )----------'
'----------------------------------------------------- 2005.03.04 追加 ( Start )----------'
' 完納予定日・便算出方法
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 16
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_KYmd(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_KYmd(0).Tag)
       Call IRNInsert
    End If
'----------------------------------------------------- 2005.03.04 追加 ( End   )----------'

' かんばん設定　仕切かんばん
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 31
    If IRNRead("", 1) Then
       IRN.oyian = H1Ck_SPrint.Value & Format(iB1_SHeigh, "00")
       Call IRNUpdate
    Else
       IRN.oyian = H1Ck_SPrint.Value & Format(iB1_SHeigh, "00")
       IRN.oyian2 = "仕切かんばん"
       Call IRNInsert
    End If
    
'+-------------------------------------------------------------------   2005/11/11 Insert
'  稼働日カレンダー
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 33
    If IRNRead("", 1) Then
       IRN.oyian = Trim(H1Op_Calendar(0).Tag)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(H1Op_Calendar(0).Tag)
       IRN.oyian2 = "稼働日カレンダー"
       Call IRNInsert
    End If
'+-------------------------------------------------------------------   2005/11/11 Insert End
    With CisFun
'                                           < 2006/02/02 MNT START >
'        .INI_FileID = gIniExe & gIniName
        .INI_FileID = gIniExe & gwsIniName
'                                           < 2006/02/02 MNT  END  >
        .INI_Section = "Option"
        '自社伝票ファイル出力先
        .INI_Key = "DENPYOFILEOUT"
        .INI_String = iB1_JDenFDir
        .PutIni
        'JAMA伝票ファイル出力先
        .INI_Key = "JAMAFILEOUT"
        .INI_String = iB1_JamaFDir
        .PutIni
'----------------------------------------------------- 2005.02.21 追加 ( Start )----------'
        '自給チケットファイル出力先
        .INI_Key = "JTICKETFILEOUT"
        .INI_String = iB1_JTicketFDir
        .PutIni
        '支給チケットファイル出力先
        .INI_Key = "STICKETFILEOUT"
        .INI_String = iB1_STicketFDir
        .PutIni
'----------------------------------------------------- 2005.02.21 追加 ( End   )----------'
        '平準化月超過ファイル出力先
        .INI_Key = "HEIFILEOUT"
        .INI_String = iB1_HeiFDir
        .PutIni
        'イメージ格納先
        .INI_Key = "Image"
        .INI_String = iB1_ImageDir
        .PutIni
    ' 接続機器
        For gInt = 0 To 6
'                                           < 2006/02/02 ADD START >
            .INI_FileID = gIniExe & gIniName
'                                           < 2006/02/02 ADD  END  >
            .INI_Section = "Apparatus"
            .INI_Key = gInt + 1
            .INI_String = Connection(gInt)
            .PutIni
            
'                                           < 2006/02/02 ADD START >
            .INI_FileID = gIniExe & gwsIniName
'                                           < 2006/02/02 ADD  END  >
            .INI_Section = Connection(gInt)
            .INI_Key = "Use"
            Select Case gInt
                Case 0: gStr = H1Op_KiKi1(0).Tag
                Case 1: gStr = H1Op_KiKi2(0).Tag
                Case 2: gStr = H1Op_KiKi3(0).Tag
                Case 3: gStr = H1Op_KiKi4(0).Tag
                Case 4: gStr = H1Op_KiKi5(0).Tag
                Case 5: gStr = H1Op_KiKi6(0).Tag
                Case 6: gStr = H1Op_KiKi7(0).Tag
            End Select
            If gStr = "0" Then
               .INI_String = "Yes"
            Else
               .INI_String = "No"
            End If
            .PutIni
        Next gInt
        ' 変換元フォルダ
        .INI_Section = "TextConvert"
        .INI_Key = "InputDir"
        .INI_String = iB1_TextDir
        .PutIni
        .INI_Key = "LogDelDay"
        .INI_String = CisFun.Val2(iB1_LogDel)
        .PutIni
        .INI_Key = "Jisya"
        If H1Ck_Jisya.Value = 1 Then
           .INI_String = "Yes"
        Else
           .INI_String = "No"
        End If
        .PutIni
        ' かんばん印刷先
        .INI_Section = "かんばん印刷先"
        .INI_Key = CisFun.WSID
        .INI_String = RTrim(lb_KPrinter)
        .PutIni

'-----     -----     -----     -----     -----     -----     < 2006/02/22 ADD START >
'
'自給チケット出力先
'
        .INI_Section = "自給チケット印刷先"
        .INI_Key = CisFun.WSID
        .INI_String = RTrim(lb_CPrinter)
        .PutIni
'
'-----     -----     -----     -----     -----     -----     < 2006/02/22 ADD  END  >

'                                           < 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gIniName
'                                           < 2006/02/02 ADD  END  >
    End With
'------------------------------------------------------ 2006/03/16 Insert
'--  ｵﾌﾟｼｮﾝ設定の内容によりﾒﾆｭｰ表示･非表示の切替を行う
    ' 発注指示(時間) or 発注指示(発注ﾃｰﾌﾞﾙ)
    If H1Op_KPool(1).Value Then     ' 即時発注は非表示にする
       Call IrnPut("MenuPXH0030", 0, "1", "かんばん指示(時間)")
       Call IrnPut("MenuPXH0060", 0, "1", "かんばん指示(ﾃｰﾌﾞﾙ)")
       Call IrnPut("MenuPKN0120", 0, "1", "発注テーブル メンテ")
    Else
       If H1Op_NYmdBin(0).Value Then
          Call IrnPut("MenuPXH0030", 0, "0", "かんばん指示(時間)")
          Call IrnPut("MenuPXH0060", 0, "1", "かんばん指示(ﾃｰﾌﾞﾙ)")
          Call IrnPut("MenuPKN0120", 0, "1", "発注テーブル メンテ")
       Else
          Call IrnPut("MenuPXH0030", 0, "1", "かんばん指示(時間)")
          Call IrnPut("MenuPXH0060", 0, "0", "かんばん指示(ﾃｰﾌﾞﾙ)")
          Call IrnPut("MenuPKN0120", 0, "0", "発注テーブル メンテ")
       End If
    End If
'------------------------------------------------------ 2006/03/16 Insert End
        
' 更新確認
    MsgBox " " & vbCrLf & "         システムオプション設定更新終了 ！！        " & vbCrLf & " ", _
            vbOKOnly + vbInformation, "オプション設定更"
    B1_Area1.Enabled = True
    
    On Error GoTo 0
    
'    Call CisDB.DBDISConnect
    
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'------------------------------------------------------ 2006/03/16 Insert
Private Sub IrnPut(bk As String, rno As Long, nai1 As String, nai2 As String)
    Call ItemsClearIRN
    IRN.bk = bk: IRN.rno = rno
    If IRNRead("", 1) Then
       IRN.oyian = nai1
       IRN.oyian2 = nai2
       Call IRNUpdate
    Else
       IRN.oyian = nai1
       IRN.oyian2 = nai2
       Call IRNInsert
    End If
End Sub
'------------------------------------------------------ 2006/03/16 Insert End
''''+----------------------+
''''+    パスワード変更     +
''''+----------------------+
'''Private Sub cmdPass_Click()
'''    frmssap.Show vbModal
'''    Unload frmssap
'''    Set frmssap = Nothing
'''End Sub
