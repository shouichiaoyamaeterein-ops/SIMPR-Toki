VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CSA0100 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "生産負荷問合せ"
   ClientHeight    =   10980
   ClientLeft      =   1740
   ClientTop       =   705
   ClientWidth     =   15315
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10980
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin VB.PictureBox B1_Area1 
      Height          =   7950
      Left            =   52
      ScaleHeight     =   7890
      ScaleWidth      =   15150
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2745
      Width           =   15210
      Begin VB.PictureBox B2_Area1 
         BackColor       =   &H00FF0000&
         Height          =   3780
         Left            =   2085
         ScaleHeight     =   3720
         ScaleWidth      =   11085
         TabIndex        =   16
         Top             =   1875
         Width           =   11145
         Begin Cis3D_v60.CIS3D Title 
            Height          =   330
            Left            =   90
            Top             =   45
            Width           =   10530
            _ExtentX        =   18574
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "詳　細"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   810
            Left            =   75
            Top             =   405
            Width           =   10920
            _ExtentX        =   19262
            _ExtentY        =   1429
            BackColor       =   16761087
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   720
               Index           =   0
               Left            =   45
               Top             =   45
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   1270
               ForeColor       =   16711680
               Caption         =   "部門"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionX      =   60
               cPositionY      =   20
               Begin Cis3D_v60.CIS3D B2lb_Bumon 
                  Height          =   360
                  Left            =   60
                  Top             =   300
                  Width           =   2205
                  _ExtentX        =   3889
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "200：プレス"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
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
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   720
               Index           =   1
               Left            =   2370
               Top             =   45
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   1270
               ForeColor       =   16711680
               Caption         =   "機械"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionX      =   60
               cPositionY      =   20
               Begin Cis3D_v60.CIS3D B2lb_Kikai 
                  Height          =   360
                  Left            =   60
                  Top             =   300
                  Width           =   2205
                  _ExtentX        =   3889
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "001：300tプレス"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
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
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   720
               Index           =   2
               Left            =   4695
               Top             =   45
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   1270
               ForeColor       =   16711680
               Caption         =   "ﾛﾎﾞｯﾄ区"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionX      =   60
               cPositionY      =   20
               Begin Cis3D_v60.CIS3D B2lb_Robot 
                  Height          =   360
                  Left            =   60
                  Top             =   300
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "ロボット"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
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
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   720
               Index           =   3
               Left            =   5865
               Top             =   45
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   1270
               ForeColor       =   16711680
               Caption         =   "機械№"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionX      =   60
               cPositionY      =   20
               Begin Cis3D_v60.CIS3D B2lb_KikaiNo 
                  Height          =   360
                  Left            =   75
                  Top             =   285
                  Width           =   2205
                  _ExtentX        =   3889
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "002：P55"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
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
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   720
               Index           =   4
               Left            =   8235
               Top             =   45
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   1270
               ForeColor       =   16711680
               Caption         =   "データ区分"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionX      =   60
               cPositionY      =   20
               Begin Cis3D_v60.CIS3D B2lb_Kubun 
                  Height          =   360
                  Left            =   75
                  Top             =   300
                  Width           =   1230
                  _ExtentX        =   2170
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "実績"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
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
                  cAlingnment     =   7
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   720
               Index           =   5
               Left            =   9615
               Top             =   45
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   1270
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionX      =   60
               cPositionY      =   20
            End
         End
         Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
            Bindings        =   "CSA0100.frx":0000
            Height          =   2385
            Left            =   75
            OleObjectBlob   =   "CSA0100.frx":0015
            TabIndex        =   17
            Top             =   1260
            Width           =   10920
         End
         Begin Cis3D_v60.CIS3D PB_Close 
            Height          =   330
            Left            =   10665
            Top             =   45
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "×"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cButton         =   -1  'True
            cPositionX      =   20
            cPositionY      =   20
         End
      End
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CSA0100.frx":13C66
         Height          =   7770
         Left            =   45
         OleObjectBlob   =   "CSA0100.frx":13C7B
         TabIndex        =   11
         Top             =   60
         Width           =   15045
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1095
      Left            =   60
      ScaleHeight     =   1035
      ScaleWidth      =   13815
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1395
      Width           =   13875
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   930
         Left            =   10440
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1640
         ForeColor       =   16711680
         Caption         =   "工数選択"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin VB.OptionButton H1Op_Kosu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "目標工数"
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
            Index           =   1
            Left            =   150
            TabIndex        =   6
            Top             =   645
            Width           =   1365
         End
         Begin VB.OptionButton H1Op_Kosu 
            BackColor       =   &H00C0FFC0&
            Caption         =   "基準工数"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   375
            Value           =   -1  'True
            Width           =   1350
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   930
         Index           =   3
         Left            =   3300
         Top             =   60
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   1640
         ForeColor       =   16711680
         Caption         =   "部   門"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionX      =   50
         cPositionY      =   75
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   120
            TabIndex        =   3
            Top             =   420
            Width           =   990
            _ExtentX        =   1746
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
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   345
            Left            =   1110
            Top             =   420
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D H1_DateTitle 
         Height          =   930
         Left            =   75
         Top             =   60
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   1640
         ForeColor       =   16711680
         Caption         =   "日付"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   100
         Begin Cis3D_v60.CIS3D PB_SDate 
            Height          =   345
            Left            =   810
            Top             =   105
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   609
            Caption         =   "開始"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D PB_EDate 
            Height          =   345
            Left            =   810
            Top             =   465
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   609
            Caption         =   "終了"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_SDate 
            Height          =   345
            Left            =   1290
            TabIndex        =   1
            Top             =   105
            Width           =   1830
            _ExtentX        =   2990
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Object.Height          =   345
            Object.Width           =   1830
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_EDate 
            Height          =   345
            Left            =   1290
            TabIndex        =   2
            Top             =   465
            Width           =   1830
            _ExtentX        =   2990
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Object.Height          =   345
            Object.Width           =   1830
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   930
         Left            =   7035
         Top             =   60
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   1640
         ForeColor       =   16711680
         Caption         =   "機　械"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Kikai 
            Height          =   375
            Left            =   135
            TabIndex        =   4
            Top             =   405
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   661
            cWildeCard      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXX"
            MaxLength       =   4
            IMEMode         =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_Kikainm 
            Height          =   375
            Left            =   780
            Top             =   390
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   661
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   930
         Left            =   12090
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1640
         ForeColor       =   16711680
         Caption         =   "表示区分"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin VB.OptionButton H1Op_DspKbn 
            BackColor       =   &H00C0FFC0&
            Caption         =   "プレス"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   7
            Top             =   375
            Value           =   -1  'True
            Width           =   1350
         End
         Begin VB.OptionButton H1Op_DspKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "スポット"
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
            Index           =   1
            Left            =   150
            TabIndex        =   8
            Top             =   645
            Width           =   1365
         End
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   705
      Left            =   11745
      ScaleHeight     =   645
      ScaleWidth      =   3420
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   450
      Width           =   3480
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   2550
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
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
         Height          =   585
         Left            =   1710
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
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
         Height          =   585
         Left            =   870
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   30
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   11
         BtnCaption      =   "検索"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   33023
         BtnShadowColor  =   16777215
         PFCaption       =   "( F7 )"
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
         PFBackColor     =   33023
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D31 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 生産負荷問合せ 】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   1
      Begin VB.PictureBox Back_Over 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5220
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Width           =   1995
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "重複"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   780
         End
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0FFC0&
            Caption         =   "単一"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   1020
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   780
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   7500
         Top             =   45
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         BackColor       =   14737632
         Caption         =   "抽出件数[              ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cPositionX      =   50
         Begin Cis3D_v60.CIS3D lb_Kensu 
            Height          =   285
            Left            =   1095
            Top             =   15
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "##,### / ##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   14737632
            cBoderColor2    =   14737632
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14415
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
         BackColor       =   14737632
         ForeColor       =   16711680
         Caption         =   "SIMPR-A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8454143
         cFont3DColor2   =   16777215
         cFont3D         =   1
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   315
         Left            =   11925
         Top             =   30
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         cBoderStyle     =   1
         cPositionX      =   50
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   0
      Top             =   495
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
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
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1230
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   510
      Top             =   0
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Text            =   "D"
      Top             =   60
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   405
      Left            =   2160
      Top             =   495
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CSA0100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   生産負荷問合せ
'**       フォームID    :   CSA0100
'**       処理概要      :
'**
'**       作  成  日    :   2010/01/26  By CIS
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean          ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2       ' 入力域切替
    Dim CisDBGrid1              As New CisDBGrid    ' TDBGridPro60ｸﾗｽ定義
    Dim CisVsGrid1              As New CisVsGrid3
'   背景色
    Dim mGotColor               As Long
    Dim mLostColor              As Long
'   退避
    Dim SaveRow                 As Long
    Dim DirWork                 As String
    
    Dim DragFlg                 As Boolean
    Dim DragX                   As Single
    Dim DragY                   As Single
    
    Dim LeftSA                  As Single
    Dim TopSA                   As Single
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub form_load()
'   #-------------------#
'   # 二 重 起 動 防 止  #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # 初 期 内 容 取 得  #
'   #-------------------#
    If Not IniGet Then End
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld

    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If

    Dummy.Left = -1000
    Dummy.Enabled = False

    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止

    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
'+-----------------------+
'+ Layファイル 初期設定  +
'+-----------------------+
    DirWork = gIniExe & CisFun.WSID & "A0100.Lay"
    If Dir(DirWork) = "" Then
        With CisFun
            .FE_FullPath = gIniExe & "A0100.Lay"
            .FE_CpySak = gIniExe & CisFun.WSID & "A0100.Lay"
            .FCpy
        End With
    End If

'+---------------------+
'+ TrueGrid60 初期設定  +
'+---------------------+
    With CisDBGrid1
        Set .DBGridObj = DBGrid1
        .FrmCaption = Me.Caption
        .MaxSort = 7                                  ' ﾚｲｱｳﾄｿｰﾄ最大数
        .LayOut = True                                ' ﾚｲｱｳﾄ変更処理 True:有り,False:無し
        .LayOutMax = 2                                ' ﾚｲｱｳﾄ情報最大数
        .LayOutColMax = 30                            ' ﾚｲｱｳﾄ表示可能項目最大数
        .LayOutPath = gIniExe                         ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ場所
        .LayOutName = CisFun.WSID & "A0100.Lay"       ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ名
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        '.CsvName = Trim(Me.Caption) & ".csv"
        '.CsvRegKey = gRegKey
        '.CsvValue = "CSA0100"
        .GridInitGet = False                 ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ
        .GridProcGet = False
        .Init
    End With
    Call GridInitSet1

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    mGotColor = H1Op_Kosu(0).BackColor
    mLostColor = H1Op_Kosu(1).BackColor

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

' Head部初期化

    H1Op_Kosu(0).Value = True: H1Op_Kosu(0).BackColor = mGotColor
    Op_Overlap(0).Value = True: Op_Overlap(0).BackColor = mGotColor

    Call HeadBodyClear("H")

' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    Adodc2.ConnectionString = CisDB.ConnectString

    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False

    Select Case Key_Code
           'Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           'Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
           'Case vbKeyF10:       If PB_Csv.Visible Then Call PB_Csv_Click    '【出力】
           'Case vbKeyF11:       If PB_Set.Visible Then Call PB_Set_Click    '【設定】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select

    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'+---------------------+
'+   ﾃﾞｰﾀﾍﾞｰｽ切断
'+---------------------+
'   #### ﾃﾞｰﾀﾍﾞｰｽ切断 ####
    Call CisDB.DBDISConnect
'   #### ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去 ####
    Call CisFrm.UnLoadDisp
    End
End Sub

Private Sub PB_Close_Click()
    Call PB_CAN_Click
End Sub
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
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        Call HeadBodyClear("H")
        'H1Cmb_UKbn.SetFocus
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call DBGridClr1
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B2" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        DBGrid1.SetFocus
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub

'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_SDate" Then iH1_SDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EDate" Then iH1_EDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Kikai" Then Call Look_Kikai: Exit Sub
End Sub
'+-----------------------------+
'+        部門検索             +
'+-----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 7

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+------------------------------+
'+      機械(名称マスタ) 検索   +
'+------------------------------+
Private Sub Look_Kikai()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "機械"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iH1_Kikai = RV_Code
        Call iH1_Kikai_LostFocus
    End If
End Sub
'+-----------------------------+
'+  iH2_Kikai
'+-----------------------------+
Private Sub iH1_Kikai_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Kikai_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    H1lb_Kikainm = ""
    If MeisyoGet("機械", RTrim(iH1_Kikai)) Then
        H1lb_Kikainm = RTrim(SYM_Meisyo)
    End If
End Sub
'+-----------------------------+
'+  iH1_Torcd
'+-----------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_Tornm = ""
    If TorNmGet(iH1_Torcd, 7) Then
        H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_SDate
'+-----------------------------+
Private Sub iH1_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_EDate
'+-----------------------------+
Private Sub iH1_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_EDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+---------------------------+
'+    重複表示　ﾁｪｯｸﾎﾞｯｸｽ     +
'+---------------------------+
Private Sub Op_Overlap_GotFocus(Index As Integer)
    For Each gObj In Op_Overlap
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub Op_Overlap_Click(Index As Integer)
    If Not FormAct Then Exit Sub

    Call CisFrm.MousePT(11)

    Call CisDBGrid1.MargeDisp(Op_Overlap(1).Value)

'''    If ProcHB = "H1" Then
'''        H1Cmb_UKbn.SetFocus
'''    Else
'''        DBGrid1.SetFocus
'''    End If
    
'''    DBGrid1.SetFocus

    Call CisFrm.MousePT(1)
End Sub
'+----------------------------------+
'+      H1Op_DspKbn
'+----------------------------------+
Private Sub H1Op_DspKbn_GotFocus(Index As Integer)
    For Each gObj In H1Op_DspKbn
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'+----------------------------+
'+    表示切替ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　     +
'+----------------------------+
Private Sub H1Op_DspKbn_Click(Index As Integer)
    Call CisFrm.MousePT(11)

    CisDBGrid1.LayOutNow = Index + 1
    CisDBGrid1.LayOutDisp

    If ProcHB = "B1" Then
        Call DBInput("R")
    End If

    Call CisFrm.MousePT(1)
End Sub
'+----------------------------------+
'+      H1Op_Kosu
'+----------------------------------+
Private Sub H1Op_Kosu_GotFocus(Index As Integer)
    For Each gObj In H1Op_Kosu
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理        *
'****************************
Private Function DispChange(Pro As String)
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    Dim B2Mode      As Boolean
    Dim B2Color     As Long

    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    B2Mode = False:     B2Color = gPLostSel

'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus

    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel: PB_ENT.BtnCaption = "入力"
        Case "B1":      B1Mode = True:  B1Color = gPGotSel: PB_ENT.BtnCaption = "詳細"
        Case "B2":      B2Mode = True:  B2Color = gPGotSel
    End Select

    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode:  B2_Area1.BackColor = B2Color
    Select Case Pro
        Case "B2":  B1_Area1.Enabled = B2Mode:  B1_Area1.BackColor = B2Color
        Case Else:  B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color
    End Select

    PB_Look.Visible = H1Mode
    PB_END.Visible = H1Mode
    B2_Area1.Visible = B2Mode
    PB_ENT.Visible = H1Mode Or B1Mode
    PB_CAN.Visible = B1Mode Or B2Mode

'   ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False

End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False: .BackColor = mLostColor
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Cmb_*" Then .ListIndex = 0
        End With
    Next gObj
    If HeadBody = "H" Then
        H1Op_DspKbn(0).Value = True
        H1Op_DspKbn(0).BackColor = mGotColor
        H1Op_Kosu(0).Value = True
        H1Op_Kosu(0).BackColor = mGotColor
        'H1_DateTitle.Caption = "起算日" & vbCr & "　・" & vbCr & "打切日"
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        DBGrid1.SetFocus
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        Call BodySet
        ProcHB = "B2"
        Call DispChange(ProcHB)
        DBGrid1.SetFocus
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

'   日付範囲（開始）
    If Trim(iH1_SDate) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始）を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If
    
    If Trim(iH1_SDate.cYear) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始 - 年）を入力して下さい          "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If
    
    If Trim(iH1_SDate.cMonth) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始 - 月）を入力して下さい          "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If
    
    If Trim(iH1_SDate.cDay) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始 - 日）を入力して下さい          "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If
    
    If iH1_SDate.cChkResult = False Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始）を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If

    '日付範囲（終了）
    If Not Trim(iH1_EDate) = "" Then
    
        If Trim(iH1_EDate.cYear) = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲（終了 - 年）を入力して下さい          "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
        
        If Trim(iH1_EDate.cMonth) = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲（終了 - 月）を入力して下さい          "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
        
        If Trim(iH1_SDate.cDay) = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲（終了 - 日）を入力して下さい          "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
        
        If iH1_EDate.cChkResult = False Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲（終了）を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
        
        If iH1_SDate > iH1_EDate Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲の大小関係が正しくありません。            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
    
    End If

'   部門
    If RTrim(iH1_Torcd) <> "" Then
        If Not TorNmGet(RTrim(iH1_Torcd), 6) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    部　門    未 登 録            "
                .MB_MSG(4) = "     【 取引先マスタ 】           "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Torcd.SetFocus
           Exit Function
        End If
    End If
    
' 機械
    If RTrim(iH1_Kikai) <> "" Then
        If Not MeisyoGet("機械", RTrim(iH1_Kikai)) Then
            With CisFun
                 .MB_Lines = 5
                 .MB_MSG(2) = "      機 械 未 登 録                "
                 .MB_MSG(4) = "      【 名称マスタ 】              "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Kikai.SetFocus
           Exit Function
        End If
    End If
    
    

    If Not DBInput Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

    Head1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示     *
'*                                    *
'**************************************
Private Function DBInput(Optional InType As String = "") As Boolean
    Dim DateStr     As String
    Dim DateSt      As String
    Dim DateEd      As String
    Dim BinSt       As String
    Dim BinEd       As String

    DBInput = True

    Call CisFrm.MousePT(11)

    On Error GoTo DBInput_Err:

' 抽出中メッセージ表示
    'PNL_Wait.Visible = True
'    PNL_PF.Enabled = False
    DoEvents
    '生産負荷問合せワーク削除
    gSL_Select = ""
    gSL_Select = gSL_Select & "Delete from 生産負荷問合せワーク " & Chr(13)
    gSL_Select = gSL_Select & " Where 作成端末 = HOST_NAME()" & Chr(13)
    
    With CisDB
        '.ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    
    '生産負荷問合せワーク作成
    gSL_Select = ""
    gSL_Select = gSL_Select & " Insert Into 生産負荷問合せワーク " & Chr(13)
    gSL_Select = gSL_Select & " Select distinct DT.*" & Chr(13)
    gSL_Select = gSL_Select & "   from (" & Chr(13)
    gSL_Select = gSL_Select & "         Select distinct HKM.取引先,HKM.機械,HKM.ロボット区分,HKM.機械NO,HKM.工程,'' 種別,HKM.工程数,HOST_NAME() 作成端末" & Chr(13)
    gSL_Select = gSL_Select & "           From 発注内示テーブル HNT" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer join 品番工順マスタ HKM" & Chr(13)
    gSL_Select = gSL_Select & "             On HKM.品番   = HNT.品番" & Chr(13)
    gSL_Select = gSL_Select & "            AND HKM.取引先 = HNT.手配先" & Chr(13)
    gSL_Select = gSL_Select & "            AND HKM.受入   = HNT.受入" & Chr(13)
    gSL_Select = gSL_Select & "          Where HNT.内示年月 = '" & Mid(iH1_SDate, 1, 6) & "'" & Chr(13)
    If RTrim(iH1_EDate) = "" Then
        gSL_Select = gSL_Select & "        And HNT.納期     = '" & iH1_SDate & "'" & Chr(13)
    Else
        gSL_Select = gSL_Select & "        And HNT.納期 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
    End If
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "        And HNT.手配先   = '" & iH1_Torcd & "'" & Chr(13)
    End If
    If RTrim(iH1_Kikai) <> "" Then
        gSL_Select = gSL_Select & "        And HKM.機械     = '" & iH1_Kikai & "'" & Chr(13)
    End If
    
    gSL_Select = gSL_Select & "         Union " & Chr(13)
    
    gSL_Select = gSL_Select & "         Select distinct HKM.取引先,HKM.機械,HKM.ロボット区分,HKM.機械NO,HKM.工程,'' 種別,HKM.工程数,HOST_NAME() 作成端末" & Chr(13)
    gSL_Select = gSL_Select & "           From 実績管理テーブル JKT" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer join 品番工順マスタ HKM" & Chr(13)
    gSL_Select = gSL_Select & "             On HKM.品番   = JKT.品番" & Chr(13)
    gSL_Select = gSL_Select & "            AND HKM.取引先 = JKT.実績先" & Chr(13)
    gSL_Select = gSL_Select & "            AND HKM.受入   = JKT.受入" & Chr(13)
    gSL_Select = gSL_Select & "            AND HKM.工順   = JKT.工順" & Chr(13)
    If RTrim(iH1_EDate) = "" Then
        gSL_Select = gSL_Select & "          Where JKT.実績日 = '" & iH1_SDate & "'" & Chr(13)
    Else
        gSL_Select = gSL_Select & "          Where JKT.実績日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
    End If
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "            And JKT.実績先 = '" & iH1_Torcd & "'" & Chr(13)
    End If
    If RTrim(iH1_Kikai) <> "" Then
        gSL_Select = gSL_Select & "            And HKM.機械   = '" & iH1_Kikai & "'" & Chr(13)
    End If
    
    gSL_Select = gSL_Select & "         Union " & Chr(13)
    
    gSL_Select = gSL_Select & "         Select distinct HKM.取引先,HKM.機械,HKM.ロボット区分,HKM.機械NO,HKM.工程,'' 種別,HKM.工程数,HOST_NAME() 作成端末" & Chr(13)
    gSL_Select = gSL_Select & "           From 作業日報テーブル SNT" & Chr(13)
    gSL_Select = gSL_Select & "           Left Outer join 品番工順マスタ HKM" & Chr(13)
    gSL_Select = gSL_Select & "             On HKM.品番   = SNT.品番" & Chr(13)
    gSL_Select = gSL_Select & "            AND HKM.取引先 = SNT.部門" & Chr(13)
    gSL_Select = gSL_Select & "            AND HKM.受入   = SNT.受入" & Chr(13)
    gSL_Select = gSL_Select & "            AND HKM.工順   = SNT.工順" & Chr(13)
    If RTrim(iH1_EDate) = "" Then
        gSL_Select = gSL_Select & "          Where SNT.作業日 = '" & iH1_SDate & "'" & Chr(13)
    Else
        gSL_Select = gSL_Select & "          Where SNT.作業日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
    End If
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "            And SNT.部門   = '" & iH1_Torcd & "'" & Chr(13)
    End If
    If RTrim(iH1_Kikai) <> "" Then
        gSL_Select = gSL_Select & "            And HKM.機械   = '" & iH1_Kikai & "'" & Chr(13)
    End If
    gSL_Select = gSL_Select & "        ) DT" & Chr(13)
    gSL_Select = gSL_Select & "  where IsNull(DT.取引先,'') <> ''" & Chr(13)
    gSL_Select = gSL_Select & "  Order by DT.取引先,DT.機械,DT.機械NO,DT.種別"
         
    With CisDB
        '.ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
        
    '生産負荷問合せワークUpdate '種別を1にする
    gSL_Select = ""
    gSL_Select = gSL_Select & " Update 生産負荷問合せワーク Set " & Chr(13)
    gSL_Select = gSL_Select & "   種別 = '1'"
    gSL_Select = gSL_Select & "  Where 作成端末 = HOST_NAME()"
    With CisDB
        '.ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
           
    '生産負荷問合せワークInsert 種別 = 2 を作成
    gSL_Select = ""
    gSL_Select = gSL_Select & " Insert Into 生産負荷問合せワーク " & Chr(13)
    gSL_Select = gSL_Select & " Select 部門,機械,ロボット区分,機械NO,工程,'2',工程数,作成端末" & Chr(13)
    gSL_Select = gSL_Select & "   From 生産負荷問合せワーク" & Chr(13)
    gSL_Select = gSL_Select & "  Where 種別 = 1" & Chr(13)
    gSL_Select = gSL_Select & "    And 作成端末 = HOST_NAME()"
    With CisDB
        '.ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

    '生産負荷問合せワークInsert 種別 = 3 を作成
    gSL_Select = ""
    gSL_Select = gSL_Select & " Insert Into 生産負荷問合せワーク " & Chr(13)
    gSL_Select = gSL_Select & " Select 部門,機械,ロボット区分,機械NO,工程,'3',工程数,作成端末" & Chr(13)
    gSL_Select = gSL_Select & "   From 生産負荷問合せワーク" & Chr(13)
    gSL_Select = gSL_Select & "  Where 種別 = 1" & Chr(13)
    gSL_Select = gSL_Select & "    And 作成端末 = HOST_NAME()"
    With CisDB
        '.ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    
    '表示データ作成
    gSL_Select = ""
    gSL_Select = gSL_Select & " Select" & Chr(13)
    gSL_Select = gSL_Select & "        Case When ISNULL(STW.部門,'') = '' then ''" & Chr(13)
    gSL_Select = gSL_Select & "             Else STW.部門 + ':' + ISNULL(THM.略称,'') END 部門," & Chr(13)
    gSL_Select = gSL_Select & "        Case When ISNULL(STW.機械,'') = '' then ''" & Chr(13)
    gSL_Select = gSL_Select & "             Else STW.機械 + ':' + ISNULL(MS.値名称,'') END 機械," & Chr(13)
    gSL_Select = gSL_Select & "        Case When STW.ロボット区分 = 0 then 'ﾛﾎﾞｯﾄ'" & Chr(13)
    gSL_Select = gSL_Select & "             Else '単発' END ロボット区分," & Chr(13)
    gSL_Select = gSL_Select & "        Case When ISNULL(STW.機械NO,'') = '' then ''" & Chr(13)
    gSL_Select = gSL_Select & "             Else STW.機械NO + ':' + ISNULL(MT.値名称,'') END 機械NO," & Chr(13)
    gSL_Select = gSL_Select & "        Case When ISNULL(STW.工程,'') = '' then ''" & Chr(13)
    gSL_Select = gSL_Select & "             Else STW.工程 + ':' + ISNULL(MY.値名称,'') END 工程," & Chr(13)
    gSL_Select = gSL_Select & "        Case When STW.種別 = 1 then '内示'" & Chr(13)
    gSL_Select = gSL_Select & "             When STW.種別 = 2 then '出荷'" & Chr(13)
    gSL_Select = gSL_Select & "             When STW.種別 = 3 then '実績' End 種別," & Chr(13)
    gSL_Select = gSL_Select & "        Case When STW.種別 = 1 then HNT.内示数量" & Chr(13)
    gSL_Select = gSL_Select & "             When STW.種別 = 2 then JKT.実績数量" & Chr(13)
    gSL_Select = gSL_Select & "             When STW.種別 = 3 then SNT.日報数量 END 数量," & Chr(13)
    gSL_Select = gSL_Select & "        Case When STW.種別 = 1 then HNT.内示表示工程数" & Chr(13)
    gSL_Select = gSL_Select & "             When STW.種別 = 2 then JKT.実績表示工程数" & Chr(13)
    gSL_Select = gSL_Select & "             When STW.種別 = 3 then SNT.日報表示工程数 END 表示工程数," & Chr(13)
    If H1Op_Kosu(0).Value Then
        gSL_Select = gSL_Select & "        Case When STW.種別 = 1 then HNT.内示基準時間" & Chr(13)
        gSL_Select = gSL_Select & "             When STW.種別 = 2 then JKT.実績基準時間" & Chr(13)
        gSL_Select = gSL_Select & "             When STW.種別 = 3 then SNT.日報基準時間 END 基準時間," & Chr(13)
    Else
        gSL_Select = gSL_Select & "        Case When STW.種別 = 1 then HNT.内示目標時間" & Chr(13)
        gSL_Select = gSL_Select & "             When STW.種別 = 2 then JKT.実績目標時間" & Chr(13)
        gSL_Select = gSL_Select & "             When STW.種別 = 3 then SNT.日報目標時間 END 基準時間," & Chr(13)
    End If
    gSL_Select = gSL_Select & "        STW.部門 ワーク部門,STW.機械 ワーク機械,STW.ロボット区分 ワークロボット," & Chr(13)
    gSL_Select = gSL_Select & "        STW.機械NO ワーク機械NO,STW.工程 ワーク工程,STW.種別 ワーク種別,STW.工程数 ワーク工程数" & Chr(13)
    gSL_Select = gSL_Select & "   From 生産負荷問合せワーク STW" & Chr(13)
    gSL_Select = gSL_Select & "   Left Outer Join (" & Chr(13)
    gSL_Select = gSL_Select & "                    Select '1' 種別,HN.手配先,HK.機械,HK.ロボット区分,HK.機械NO,HK.工程,HK.工程数," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(HN.数量) 内示数量," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(HN.数量) * HK.工程数 内示表示工程数," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(HN.数量) * HK.基準工数1 内示基準時間," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(HN.数量) * HK.基準工数2 内示目標時間" & Chr(13)
    gSL_Select = gSL_Select & "                      from 発注内示テーブル HN" & Chr(13)
    gSL_Select = gSL_Select & "                      Left outer Join 品番工順マスタ HK" & Chr(13)
    gSL_Select = gSL_Select & "                        ON HK.品番   = HN.品番" & Chr(13)
    gSL_Select = gSL_Select & "                       AND HK.取引先 = HN.手配先" & Chr(13)
    gSL_Select = gSL_Select & "                       AND HK.受入   = HN.受入" & Chr(13)
    gSL_Select = gSL_Select & "                     Where HN.内示年月 = '" & Mid(iH1_SDate, 1, 6) & "'" & Chr(13)
    If RTrim(iH1_EDate) = "" Then
        gSL_Select = gSL_Select & "                   AND HN.納期     = '" & iH1_SDate & "'" & Chr(13)
    Else
        gSL_Select = gSL_Select & "                   AND HN.納期 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
    End If
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "                   AND HN.手配先   = '" & iH1_Torcd & "'" & Chr(13)
    End If
    If RTrim(iH1_Kikai) <> "" Then
        gSL_Select = gSL_Select & "                   AND HK.機械   = '" & iH1_Kikai & "'" & Chr(13)
    End If
    gSL_Select = gSL_Select & "                     Group By HN.手配先,HK.機械,HK.ロボット区分,HK.機械NO,HK.工程,HK.工程数," & Chr(13)
    gSL_Select = gSL_Select & "                              HK.基準工数1 , HK.基準工数2" & Chr(13)
    gSL_Select = gSL_Select & "                     ) HNT" & Chr(13)
    gSL_Select = gSL_Select & "     On HNT.手配先       = STW.部門" & Chr(13)
    gSL_Select = gSL_Select & "    AND HNT.機械         = STW.機械" & Chr(13)
    gSL_Select = gSL_Select & "    AND HNT.ロボット区分 = STW.ロボット区分" & Chr(13)
    gSL_Select = gSL_Select & "    AND HNT.機械NO       = STW.機械NO" & Chr(13)
    gSL_Select = gSL_Select & "    AND HNT.工程         = STW.工程" & Chr(13)
    gSL_Select = gSL_Select & "    AND HNT.工程数       = STW.工程数" & Chr(13)
    gSL_Select = gSL_Select & "   Left Outer Join (" & Chr(13)
    gSL_Select = gSL_Select & "                    Select '2' 種別,JK.実績先,HK.機械,HK.ロボット区分,HK.機械NO,HK.工程,HK.工程数," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(JK.実績数) 実績数量," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(JK.実績数) * HK.工程数 実績表示工程数," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(JK.実績数) * HK.基準工数1 実績基準時間," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(JK.実績数) * HK.基準工数2 実績目標時間" & Chr(13)
    gSL_Select = gSL_Select & "                      from 実績管理テーブル JK" & Chr(13)
    gSL_Select = gSL_Select & "                      Left outer Join 品番工順マスタ HK" & Chr(13)
    gSL_Select = gSL_Select & "                        ON HK.品番   = JK.品番" & Chr(13)
    gSL_Select = gSL_Select & "                       AND HK.取引先 = JK.実績先" & Chr(13)
    gSL_Select = gSL_Select & "                       AND HK.受入   = JK.受入" & Chr(13)
    gSL_Select = gSL_Select & "                       AND HK.工順   = JK.工順" & Chr(13)
    If RTrim(iH1_EDate) = "" Then
        gSL_Select = gSL_Select & "                  Where JK.実績日      = '" & iH1_SDate & "'" & Chr(13)
    Else
        gSL_Select = gSL_Select & "                  Where JK.実績日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
    End If
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "                    AND JK.実績先   = '" & iH1_Torcd & "'" & Chr(13)
    End If
    If RTrim(iH1_Kikai) <> "" Then
        gSL_Select = gSL_Select & "                    AND HK.機械   = '" & iH1_Kikai & "'" & Chr(13)
    End If
    gSL_Select = gSL_Select & "                      Group By JK.実績先,HK.機械,HK.ロボット区分,HK.機械NO,HK.工程,HK.工程数," & Chr(13)
    gSL_Select = gSL_Select & "                               HK.基準工数1 , HK.基準工数2" & Chr(13)
    gSL_Select = gSL_Select & "                     ) JKT" & Chr(13)
    gSL_Select = gSL_Select & "     On JKT.実績先       = STW.部門" & Chr(13)
    gSL_Select = gSL_Select & "    AND JKT.機械         = STW.機械" & Chr(13)
    gSL_Select = gSL_Select & "    AND JKT.ロボット区分 = STW.ロボット区分" & Chr(13)
    gSL_Select = gSL_Select & "    AND JKT.機械NO       = STW.機械NO" & Chr(13)
    gSL_Select = gSL_Select & "    AND JKT.工程         = STW.工程" & Chr(13)
    gSL_Select = gSL_Select & "    AND JKT.工程数       = STW.工程数" & Chr(13)
    gSL_Select = gSL_Select & "   Left Outer join (Select '3' 種別,SN.部門,HK.機械,HK.ロボット区分,HK.機械NO,HK.工程,HK.工程数," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(SN.加工数) 日報数量," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(SN.加工数) * HK.工程数 日報表示工程数," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(SN.加工数) * HK.基準工数1 日報基準時間," & Chr(13)
    gSL_Select = gSL_Select & "                           Sum(SN.加工数) * HK.基準工数2 日報目標時間" & Chr(13)
    gSL_Select = gSL_Select & "                      From 作業日報テーブル SN" & Chr(13)
    gSL_Select = gSL_Select & "                      Left outer Join 品番工順マスタ HK" & Chr(13)
    gSL_Select = gSL_Select & "                        ON HK.品番   = SN.品番" & Chr(13)
    gSL_Select = gSL_Select & "                       AND HK.取引先 = SN.部門" & Chr(13)
    gSL_Select = gSL_Select & "                       AND HK.受入   = SN.受入" & Chr(13)
    gSL_Select = gSL_Select & "                       AND HK.工順   = SN.工順" & Chr(13)
    If RTrim(iH1_EDate) = "" Then
        gSL_Select = gSL_Select & "                  Where SN.作業日      = '" & iH1_SDate & "'" & Chr(13)
    Else
        gSL_Select = gSL_Select & "                  Where SN.作業日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
    End If
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "                    AND SN.部門       = '" & iH1_Torcd & "'" & Chr(13)
    End If
    If RTrim(iH1_Kikai) <> "" Then
        gSL_Select = gSL_Select & "                    AND HK.機械       = '" & iH1_Kikai & "'" & Chr(13)
    End If
    gSL_Select = gSL_Select & "                      Group By SN.部門,HK.機械,HK.ロボット区分,HK.機械NO,HK.工程,HK.工程数," & Chr(13)
    gSL_Select = gSL_Select & "                               HK.基準工数1 , HK.基準工数2" & Chr(13)
    gSL_Select = gSL_Select & "                    ) SNT" & Chr(13)
    gSL_Select = gSL_Select & "     On SNT.部門         = STW.部門" & Chr(13)
    gSL_Select = gSL_Select & "    AND SNT.機械         = STW.機械" & Chr(13)
    gSL_Select = gSL_Select & "    AND SNT.ロボット区分 = STW.ロボット区分" & Chr(13)
    gSL_Select = gSL_Select & "    AND SNT.機械NO       = STW.機械NO" & Chr(13)
    gSL_Select = gSL_Select & "    AND SNT.工程         = STW.工程" & Chr(13)
    gSL_Select = gSL_Select & "    AND SNT.工程数       = STW.工程数" & Chr(13)
    gSL_Select = gSL_Select & "   Left Outer Join 取引先マスタ THM" & Chr(13)
    gSL_Select = gSL_Select & "     On THM.取引先CD   = STW.部門" & Chr(13)
    gSL_Select = gSL_Select & "    AND THM.取引先区分 = '1'" & Chr(13)
    gSL_Select = gSL_Select & "    AND THM.契約先区分 = '0'" & Chr(13)
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MS" & Chr(13)
    gSL_Select = gSL_Select & "     On MS.区分名称 = '機械'" & Chr(13)
    gSL_Select = gSL_Select & "    AND MS.値       = STW.機械" & Chr(13)
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MT" & Chr(13)
    gSL_Select = gSL_Select & "     On MT.区分名称 = '機械NO'" & Chr(13)
    gSL_Select = gSL_Select & "    AND MT.値       = STW.機械NO" & Chr(13)
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MY" & Chr(13)
    gSL_Select = gSL_Select & "     On MY.区分名称 = '工程管理'" & Chr(13)
    gSL_Select = gSL_Select & "    AND MY.値       = STW.工程" & Chr(13)
    gSL_Select = gSL_Select & " Where STW.作成端末 = HOST_NAME() "
    gSL_Select = gSL_Select & " Order by STW.部門,STW.機械,STW.ロボット区分,STW.機械NO,STW.工程,HNT.内示数量,STW.工程数,STW.種別" & Chr(13)

' グリッド内表示
    If Not DBGridDisp Then
        DBInput = False
        GoTo DBInput_Ed
    End If
    On Error Resume Next
' 表示データなし
    On Error GoTo 0
    If Adodc1.Recordset.EOF Then
        DBInput = False
        GoTo DBInput_Ed
    End If

' 抽出件数表示
    lb_Kensu.Visible = True
    Call KensuDisp

    On Error GoTo 0
DBInput_Ed:
    Call CisFrm.MousePT(0)

    'PNL_Wait.Visible = False
    'PNL_PF.Enabled = True
    SendKeys "{END}"
    SendKeys "{RIGHT}"
    SendKeys "{LEFT}"
    SendKeys "{HOME}"
    Exit Function
DBInput_Err:
    Call CisFrm.MousePT(0)
    CisFun.ErrorBox
    'PNL_Wait.Visible = False
    'PNL_PF.Enabled = True
End Function

'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'+==========================================================+
'+      グリッド処理(TrueDBGrid)
'+==========================================================+
'+----------------------------------------------------------+
'+   グリッド内容初期設定
'+----------------------------------------------------------+
Private Sub GridInitSet1()
' Call .GridInitSet(ﾏｰｼﾞ処理,分類処理,ﾃﾞｰﾀ連結名,印刷見出し,印刷桁数,合計印刷)
    With CisDBGrid1
         Call .GridInitSet(True, , "部門", "部門", 25)
         Call .GridInitSet(True, , "機械", "機械", 25)
         Call .GridInitSet(True, , "ロボット区分", "ﾛﾎﾞｯﾄ区", 8)
         Call .GridInitSet(True, , "機械NO", "機械№", 25, False)
         Call .GridInitSet(True, , "工程", "工程", 25, False)
         Call .GridInitSet(, , "種別", "ﾃﾞｰﾀ区分", 4, True)
         Call .GridInitSet(, , "数量", "数量", 10, True)
         Call .GridInitSet(, , "表示工程数", "工程数", 10, False)
         Call .GridInitSet(, , "基準時間", "基準時間", 10)
    End With
End Sub
'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub KensuDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    If NowCnt < 0 Then Exit Sub
    lb_Kensu.Caption = Format(NowCnt, "###,##0") & " / " & _
                       Format(MaxCnt, "###,##0")

End Sub
'+---------------------+
'+  グリッド内容表示    +
'+---------------------+
Private Function DBGridDisp()
    DBGridDisp = False

    On Error GoTo DBGridDisp_ER:

    On Error Resume Next

    Adodc1.RecordSource = gSL_Select
    DBGrid1.ReBind
    Adodc1.Refresh
    DBGrid1.Refresh

' 重複表示
    Call CisDBGrid1.MargeDisp(Op_Overlap(1).Value)
    Call KensuDisp
    DBGrid1.SetFocus
    SendKeys "{HOME}"

    On Error GoTo 0

    DBGridDisp = True
    Exit Function
DBGridDisp_ER:
    If Err <> 40002 Then        ' RDO タイムアウト
       CisFun.ErrorBox
    End If
End Function
'+-----------------------+
'+  グリッド内容初期化    +
'+-----------------------+
Private Sub DBGridClr1()

    Call CisDBGrid1.SplitsClr("", 0)   ' 分割ﾍﾟｲﾝ解除

    Adodc1.Recordset.Close
    DBGrid1.Close True

    lb_Kensu.Caption = ""

End Sub
'+----------------------------------------------------------+
'+ キー入力にて行移動 《 KeyUp イベント   》                  +
'+----------------------------------------------------------+
Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or _
       KeyCode = vbKeyDown Or _
       KeyCode = vbKeyPageUp Or _
       KeyCode = vbKeyPageDown Then
        Call KensuDisp
        If B2_Area1.Visible Then
            Call BodySet
        End If
    End If
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 ColMove イベント   》            +
'+----------------------------------------------------------+
Private Sub DBGrid1_ColMove(ByVal Position As Integer, Cancel As Integer)
    CisDBGrid1.ColMoveFlg = True
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 Paint イベント   》              +
'+----------------------------------------------------------+
Private Sub DBGrid1_Paint()
    If Not FormAct Then Exit Sub
    If CisDBGrid1.ColMoveFlg Then  ' 再表示(分類)
        CisDBGrid1.ColMoveFlg = False
        Call DBInput
    End If
End Sub
'+----------------------------------------------------------+
'+ 明細画面表示(B2_AREA1)
'+----------------------------------------------------------+
Private Sub DBGrid1_DblClick()
    If Not B2_Area1.Visible Then
        Call ReturnPress
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseDown  イベント   》               +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If DBGrid1.RowContaining(Y) = -1 And _
        DBGrid1.ColContaining(X) = -1 Then
        Call CisDBGrid1.SplitsClr("C", 0)           ' 分割ﾍﾟｲﾝのみ解除
        Call DBInput
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then            ' 分割ﾍﾟｲﾝ表示
        If Button = 2 Then
            Call CisDBGrid1.SplitsAdd(DBGrid1.ColContaining(X))
            Call DBInput
        End If
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
    End If
    
    'If DBGrid1.RowContaining(X) < 0 Then Exit Sub

    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
    Else
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
    End If

End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call KensuDisp
End Sub
'**********************************************************
'*  明細表示
'**********************************************************
Private Sub BodySet()
    Dim ICnt        As Integer
    Dim wUKbn       As Byte
    Dim wHinbn      As String
    Dim wTorcd      As String
    Dim wKbn        As Byte
    
    Call CisFrm.MousePT(11)
    
    B2lb_Bumon = RTrim(DBGrid1.Columns(0))
    B2lb_Kikai = RTrim(DBGrid1.Columns(1))
    B2lb_Robot = RTrim(DBGrid1.Columns(2))
    B2lb_KikaiNo = RTrim(DBGrid1.Columns(3))
    B2lb_Kubun = RTrim(DBGrid1.Columns(5))
        
    wTorcd = RTrim(DBGrid1.Columns(9))
    wKbn = RTrim(DBGrid1.Columns(14))
    
    If wKbn = 1 Then '内示
        gSL_Select = ""
        gSL_Select = gSL_Select & " Select HM.表示品番,HM.品名," & Chr(13)
        gSL_Select = gSL_Select & "        Sum(HNT.数量) 数量," & Chr(13)
        gSL_Select = gSL_Select & "        Sum(HNT.数量) * HKM.工程数 工程数," & Chr(13)
        If H1Op_Kosu(0).Value Then
            gSL_Select = gSL_Select & "        Sum(HNT.数量) * HKM.基準工数1 基準時間" & Chr(13)
        Else
            gSL_Select = gSL_Select & "        Sum(HNT.数量) * HKM.基準工数2 基準時間" & Chr(13)
        End If
        gSL_Select = gSL_Select & "   From 発注内示テーブル HNT" & Chr(13)
        gSL_Select = gSL_Select & "   Left outer join 品番工順マスタ HKM" & Chr(13)
        gSL_Select = gSL_Select & "     On HKM.品番 = HNT.品番" & Chr(13)
        gSL_Select = gSL_Select & "    AND HKM.取引先 = HNT.手配先" & Chr(13)
        gSL_Select = gSL_Select & "    AND HKM.受入 = HNT.受入" & Chr(13)
        gSL_Select = gSL_Select & "   Left Outer Join 品番マスタ HM" & Chr(13)
        gSL_Select = gSL_Select & "     On HM.品番 = HNT.品番" & Chr(13)
        gSL_Select = gSL_Select & "  Where HNT.内示年月 = '" & Mid(iH1_SDate, 1, 6) & "'" & Chr(13)
        If RTrim(iH1_EDate) = "" Then
            gSL_Select = gSL_Select & "    And HNT.納期     = '" & iH1_SDate & "'" & Chr(13)
        Else
            gSL_Select = gSL_Select & "    And HNT.納期 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
        End If
        gSL_Select = gSL_Select & "    And ISNULL(HKM.品番,'') <> ''" & Chr(13)
        gSL_Select = gSL_Select & "    And HNT.手配先 = '" & wTorcd & "'" & Chr(13)
        gSL_Select = gSL_Select & "    And HKM.機械           = '" & RTrim(DBGrid1.Columns(10)) & "'" & Chr(13)
        gSL_Select = gSL_Select & "    And HKM.ロボット区分   = '" & RTrim(DBGrid1.Columns(11)) & "'" & Chr(13)
        gSL_Select = gSL_Select & "    And HKM.機械NO         = '" & RTrim(DBGrid1.Columns(12)) & "'" & Chr(13)
        gSL_Select = gSL_Select & "    And HKM.工程           = '" & RTrim(DBGrid1.Columns(13)) & "'" & Chr(13)
        gSL_Select = gSL_Select & "    And HKM.工程数         = '" & RTrim(DBGrid1.Columns(15)) & "'" & Chr(13)
        gSL_Select = gSL_Select & " Group by HM.表示品番,HM.品名,HKM.工程数,HKM.基準工数1,HKM.基準工数2" & Chr(13)
   
    Else
        If wKbn = 2 Then '出荷
            gSL_Select = ""
            gSL_Select = gSL_Select & " Select HM.表示品番,HM.品名," & Chr(13)
            gSL_Select = gSL_Select & "        Sum(JKT.実績数) 数量," & Chr(13)
            gSL_Select = gSL_Select & "        Sum(JKT.実績数) * HKM.工程数 工程数," & Chr(13)
            If H1Op_Kosu(0).Value Then
                gSL_Select = gSL_Select & "        Sum(JKT.実績数) * HKM.基準工数1 基準時間" & Chr(13)
            Else
                gSL_Select = gSL_Select & "        Sum(JKT.実績数) * HKM.基準工数2 基準時間" & Chr(13)
            End If
            gSL_Select = gSL_Select & "   From 実績管理テーブル JKT" & Chr(13)
            gSL_Select = gSL_Select & "   Left outer join 品番工順マスタ HKM" & Chr(13)
            gSL_Select = gSL_Select & "     On HKM.品番 = JKT.品番" & Chr(13)
            gSL_Select = gSL_Select & "    AND HKM.取引先 = JKT.実績先" & Chr(13)
            gSL_Select = gSL_Select & "    AND HKM.受入 = JKT.受入" & Chr(13)
            gSL_Select = gSL_Select & "    AND HKM.工順 = JKT.工順" & Chr(13)
            gSL_Select = gSL_Select & "   Left Outer Join 品番マスタ HM" & Chr(13)
            gSL_Select = gSL_Select & "     On HM.品番 = JKT.品番" & Chr(13)
            If RTrim(iH1_EDate) = "" Then
                gSL_Select = gSL_Select & "    Where JKT.実績日     = '" & iH1_SDate & "'" & Chr(13)
            Else
                gSL_Select = gSL_Select & "    Where JKT.実績日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
            End If
            gSL_Select = gSL_Select & "    And ISNULL(HKM.品番,'') <> ''"
            gSL_Select = gSL_Select & "    And JKT.実績先 = '" & wTorcd & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.機械           = '" & RTrim(DBGrid1.Columns(10)) & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.ロボット区分   = '" & RTrim(DBGrid1.Columns(11)) & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.機械NO         = '" & RTrim(DBGrid1.Columns(12)) & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.工程           = '" & RTrim(DBGrid1.Columns(13)) & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.工程数         = '" & RTrim(DBGrid1.Columns(15)) & "'" & Chr(13)
            gSL_Select = gSL_Select & " Group by HM.表示品番,HM.品名,HKM.工程数,HKM.基準工数1,HKM.基準工数2" & Chr(13)
        Else '実績
            gSL_Select = ""
            gSL_Select = gSL_Select & " Select HM.表示品番,HM.品名," & Chr(13)
            gSL_Select = gSL_Select & "        Sum(SNT.加工数) 数量," & Chr(13)
            gSL_Select = gSL_Select & "        Sum(SNT.加工数) * HKM.工程数 工程数," & Chr(13)
            If H1Op_Kosu(0).Value Then
                gSL_Select = gSL_Select & "        Sum(SNT.加工数) * HKM.基準工数1 基準時間" & Chr(13)
            Else
                gSL_Select = gSL_Select & "        Sum(SNT.加工数) * HKM.基準工数2 基準時間" & Chr(13)
            End If
            gSL_Select = gSL_Select & "   From 作業日報テーブル SNT" & Chr(13)
            gSL_Select = gSL_Select & "   Left outer join 品番工順マスタ HKM" & Chr(13)
            gSL_Select = gSL_Select & "     On HKM.品番 = SNT.品番" & Chr(13)
            gSL_Select = gSL_Select & "    AND HKM.取引先 = SNT.部門" & Chr(13)
            gSL_Select = gSL_Select & "    AND HKM.受入 = SNT.受入" & Chr(13)
            gSL_Select = gSL_Select & "    AND HKM.工順 = SNT.工順" & Chr(13)
            gSL_Select = gSL_Select & "   Left Outer Join 品番マスタ HM" & Chr(13)
            gSL_Select = gSL_Select & "     On HM.品番 = SNT.品番" & Chr(13)
            If RTrim(iH1_EDate) = "" Then
                gSL_Select = gSL_Select & "    Where SNT.作業日     = '" & iH1_SDate & "'" & Chr(13)
            Else
                gSL_Select = gSL_Select & "    Where SNT.作業日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
            End If
            gSL_Select = gSL_Select & "    And ISNULL(HKM.品番,'') <> ''" & Chr(13)
            gSL_Select = gSL_Select & "    And SNT.部門           = '" & wTorcd & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.機械           = '" & RTrim(DBGrid1.Columns(10)) & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.ロボット区分   = '" & RTrim(DBGrid1.Columns(11)) & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.機械NO         = '" & RTrim(DBGrid1.Columns(12)) & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.工程           = '" & RTrim(DBGrid1.Columns(13)) & "'" & Chr(13)
            gSL_Select = gSL_Select & "    And HKM.工程数         = '" & RTrim(DBGrid1.Columns(15)) & "'" & Chr(13)
            gSL_Select = gSL_Select & " Group by HM.表示品番,HM.品名,HKM.工程数,HKM.基準工数1,HKM.基準工数2" & Chr(13)
        
        End If
    End If
    
    Call DBGridDisp2

    Call CisFrm.MousePT(0)
End Sub
'+------------------------------+
'+  グリッド内容表示（詳細）    +
'+------------------------------+
Private Function DBGridDisp2()
    DBGridDisp2 = False

    On Error GoTo DBGridDisp2_ER:

    On Error Resume Next

    Adodc2.RecordSource = gSL_Select
    TDBGrid1.ReBind
    Adodc2.Refresh
    TDBGrid1.Refresh

    On Error GoTo 0

    DBGridDisp2 = True
    Exit Function
DBGridDisp2_ER:
    If Err <> 40002 Then        ' RDO タイムアウト
       CisFun.ErrorBox
    End If
End Function
'+--------------------------------------------+
'+  B2_Area1の移動（B2_Area1 - MouseDown時）  +
'+--------------------------------------------+
Private Sub B2_Area1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragFlg = True
    DragX = X
    DragY = Y
End Sub
Private Sub B2_Area1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim WorkX As Single
    Dim WorkY As Single
    
    If DragFlg = True Then
        With B2_Area1
        
            WorkX = .Left + X - DragX
            WorkY = .Top + Y - DragY
        
            If WorkX < 0 Then
                WorkX = 0
            Else
                If WorkX + .Width >= B1_Area1.Width Then
                    WorkX = B1_Area1.Width - .Width
                End If
            End If
    
            If WorkY < 0 Then
                WorkY = 0
            Else
                If WorkY + .Height >= B1_Area1.Height Then
                    WorkY = B1_Area1.Height - .Height
                End If
            End If
            
            .Move WorkX, WorkY

        End With
    End If
End Sub
Private Sub B2_Area1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragFlg = False
    DBGrid1.SetFocus
End Sub

'+----------------------------------------+
'+  B2_Area1の移動（Title - MouseDown時） +
'+----------------------------------------+
Private Sub Title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragFlg = True
    DragX = X
    DragY = Y
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim WorkX As Single
    Dim WorkY As Single
    
    If DragFlg = True Then
        With B2_Area1
        
            WorkX = .Left + X - DragX - Title.Left
            WorkY = .Top + Y - DragY - Title.Top
        
            If WorkX < 0 Then
                WorkX = 0
            Else
                If WorkX + .Width >= B1_Area1.Width Then
                    WorkX = B1_Area1.Width - .Width
                End If
            End If
    
            If WorkY < 0 Then
                WorkY = 0
            Else
                If WorkY + .Height >= B1_Area1.Height Then
                    WorkY = B1_Area1.Height - .Height
                End If
            End If
            
            .Move WorkX, WorkY

        End With
    End If
End Sub

Private Sub Title_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragFlg = False
    DBGrid1.SetFocus
End Sub
