VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CisYM_6.0.ocx"
Begin VB.Form CSA0070 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "納入指示確認"
   ClientHeight    =   11190
   ClientLeft      =   1785
   ClientTop       =   750
   ClientWidth     =   15420
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11190
   ScaleWidth      =   15420
   Begin Cis3D_v60.CIS3D Sel_OyaHinbn 
      Height          =   420
      Left            =   8370
      Top             =   2505
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   741
      BackColor       =   16777152
      Caption         =   "親品番：XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
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
      cAlingnment     =   1
      cBoderStyle     =   1
      cPositionX      =   60
   End
   Begin VB.PictureBox H3_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1845
      Left            =   1380
      ScaleHeight     =   1818.26
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   4110
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2955
      Width           =   4170
      Begin Cis3D_v60.CIS3D CIS3D31 
         Height          =   1680
         Left            =   30
         Top             =   60
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   2963
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   1350
            Left            =   0
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   315
            Width           =   4020
            _Version        =   196608
            _ExtentX        =   7091
            _ExtentY        =   2381
            _StockProps     =   228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "　　|親　　　品　　　番　　　　　　　　　|"
            Rows            =   5
            Cols            =   4
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D34 
            Height          =   330
            Left            =   0
            Top             =   0
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   582
            BackColor       =   16744576
            ForeColor       =   0
            Caption         =   "※　親品番選択　※"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cAlingnment     =   7
         End
      End
   End
   Begin VB.PictureBox H2_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   15
      ScaleHeight     =   960
      ScaleWidth      =   8265
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1920
      Width           =   8325
      Begin Cis3D_v60.CIS3D Back_Kobetu 
         Height          =   870
         Left            =   60
         Top             =   45
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "表示内容"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   420
            Left            =   375
            Top             =   360
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   741
            BackColor       =   14737632
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
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
            Begin Cis3D_v60.CIS3D Back_DspKosei 
               Height          =   315
               Left            =   1170
               Top             =   60
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   556
               BackColor       =   14737632
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9
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
               Begin VB.OptionButton H2Op_DspKosei 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "指定品番以下～"
                  BeginProperty Font 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   1590
                  TabIndex        =   4
                  Top             =   60
                  Width           =   1935
               End
               Begin VB.OptionButton H2Op_DspKosei 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "親品番～"
                  BeginProperty Font 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   150
                  TabIndex        =   3
                  Top             =   45
                  Value           =   -1  'True
                  Width           =   1230
               End
            End
            Begin VB.OptionButton H2Op_DspKB 
               BackColor       =   &H00E0E0E0&
               Caption         =   "詳細（工順情報）"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   5265
               TabIndex        =   5
               Top             =   75
               Width           =   1980
            End
            Begin VB.OptionButton H2Op_DspKB 
               BackColor       =   &H00C0FFC0&
               Caption         =   "構成"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   135
               TabIndex        =   2
               Top             =   75
               Value           =   -1  'True
               Width           =   975
            End
         End
      End
      Begin Cis3D_v60.CIS3D Back_TorKB 
         Height          =   870
         Left            =   60
         Top             =   45
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   1535
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   750
            Left            =   5370
            Top             =   60
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   1323
            ForeColor       =   16711680
            Caption         =   "品　　番"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
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
            cAlingnment     =   6
            cBoderStyle     =   1
            Begin CisText_V60.CisText iH2_Hinbn 
               Height          =   330
               Left            =   60
               TabIndex        =   10
               Top             =   345
               Width           =   2640
               _ExtentX        =   4657
               _ExtentY        =   582
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
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   25
               cAutoEject      =   0   'False
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   750
            Left            =   4485
            Top             =   60
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   1323
            ForeColor       =   16711680
            Caption         =   "車型"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
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
            cAlingnment     =   6
            cBoderStyle     =   1
            Begin CisText_V60.CisText iH2_Syakei 
               Height          =   360
               Left            =   60
               TabIndex        =   9
               Top             =   345
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "XXXXX"
               MaxLength       =   5
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   750
            Left            =   60
            Top             =   60
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   1323
            ForeColor       =   16711680
            Caption         =   "ﾃﾞｰﾀ有無"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
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
            cAlingnment     =   6
            cBoderStyle     =   1
            Begin VB.OptionButton H2Op_DataKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "無"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   690
               TabIndex        =   7
               Top             =   390
               Width           =   705
            End
            Begin VB.OptionButton H2Op_DataKB 
               BackColor       =   &H00C0FFC0&
               Caption         =   "有"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   45
               TabIndex        =   6
               Top             =   390
               Value           =   -1  'True
               Width           =   705
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D11 
            Height          =   750
            Left            =   1500
            Top             =   60
            Width           =   2970
            _ExtentX        =   5239
            _ExtentY        =   1323
            ForeColor       =   16711680
            Caption         =   "代 表 得 意 先"
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
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   30
            Begin CisText_V60.CisText iH2_Tokui 
               Height          =   360
               Left            =   60
               TabIndex        =   8
               Top             =   330
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   635
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
               Text            =   "XXXXXXX"
               MaxLength       =   7
            End
            Begin Cis3D_v60.CIS3D H2lb_Tokui 
               Height          =   360
               Left            =   1035
               Top             =   315
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   635
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
      End
   End
   Begin VB.PictureBox B1_Area3 
      BackColor       =   &H00FF0000&
      Height          =   1065
      Left            =   15
      ScaleHeight     =   1005
      ScaleWidth      =   11625
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3060
      Width           =   11685
      Begin Cis3D_v60.CIS3D PB_Apl 
         Height          =   450
         Index           =   1
         Left            =   30
         Top             =   60
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   794
         BackColor       =   -2147483633
         ForeColor       =   16512
         Caption         =   "処理名－１"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PB_Apl 
         Height          =   450
         Index           =   8
         Left            =   8670
         Top             =   510
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   794
         BackColor       =   -2147483633
         ForeColor       =   16512
         Caption         =   "処理名－８"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PB_Apl 
         Height          =   450
         Index           =   2
         Left            =   2910
         Top             =   60
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   794
         BackColor       =   -2147483633
         ForeColor       =   16512
         Caption         =   "処理名－２"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PB_Apl 
         Height          =   450
         Index           =   3
         Left            =   5790
         Top             =   60
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   794
         BackColor       =   -2147483633
         ForeColor       =   16512
         Caption         =   "処理名－３"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PB_Apl 
         Height          =   450
         Index           =   4
         Left            =   8670
         Top             =   60
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   794
         BackColor       =   -2147483633
         ForeColor       =   16512
         Caption         =   "処理名－４"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PB_Apl 
         Height          =   450
         Index           =   5
         Left            =   30
         Top             =   510
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   794
         BackColor       =   -2147483633
         ForeColor       =   16512
         Caption         =   "処理名－５"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PB_Apl 
         Height          =   450
         Index           =   6
         Left            =   2910
         Top             =   510
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   794
         BackColor       =   -2147483633
         ForeColor       =   16512
         Caption         =   "処理名－６"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PB_Apl 
         Height          =   450
         Index           =   7
         Left            =   5790
         Top             =   510
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   794
         BackColor       =   -2147483633
         ForeColor       =   16512
         Caption         =   "処理名－７"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cButton         =   -1  'True
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   6795
      Left            =   30
      ScaleHeight     =   6735
      ScaleWidth      =   15405
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4155
      Width           =   15465
      Begin vsFlexLib.vsFlexArray vsGrid2 
         Height          =   6600
         Left            =   45
         TabIndex        =   24
         Top             =   60
         Width           =   15255
         _Version        =   196608
         _ExtentX        =   26908
         _ExtentY        =   11642
         _StockProps     =   228
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         Rows            =   30000
         Cols            =   64
         FixedCols       =   3
         BackColor       =   12648447
         ForeColorFixed  =   16711680
         HighLight       =   2
         RowHeightMin    =   300
         AllowBigSelection=   0   'False
         BackColorAlternate=   12648447
         Begin Cis3D_v60.CIS3D Proc_MSG 
            Height          =   990
            Left            =   4785
            Top             =   2310
            Visible         =   0   'False
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   1746
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderWidth     =   5
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   765
               Index           =   1
               Left            =   105
               Top             =   105
               Width           =   5925
               _ExtentX        =   10451
               _ExtentY        =   1349
               BackColor       =   16761087
               ForeColor       =   16711680
               Caption         =   "【処理中】しばらくお待ち下さい"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐ明朝"
                  Size            =   18
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
               cBoderWidth     =   5
               cAlingnment     =   7
               cBoderStyle     =   1
            End
         End
      End
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   0
      TabIndex        =   14
      Text            =   "Dummy"
      Top             =   240
      Width           =   195
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   15
      ScaleHeight     =   750
      ScaleWidth      =   9840
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   960
      Width           =   9900
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   720
         Left            =   30
         Top             =   15
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "未来在庫計算"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   300
            Left            =   45
            Top             =   330
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            BackColor       =   14737632
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
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
            Begin VB.OptionButton H1Op_Zaiko 
               BackColor       =   &H00E0E0E0&
               Caption         =   "有"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   105
               TabIndex        =   17
               Top             =   45
               Value           =   -1  'True
               Width           =   630
            End
            Begin VB.OptionButton H1Op_Zaiko 
               BackColor       =   &H00E0E0E0&
               Caption         =   "無"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   870
               TabIndex        =   18
               Top             =   60
               Width           =   645
            End
         End
      End
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   720
         Left            =   6915
         Top             =   15
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "品  番"
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   330
            Left            =   60
            TabIndex        =   1
            Top             =   315
            Width           =   2820
            _ExtentX        =   4974
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
            Text            =   "6168312050"
            MaxLength       =   25
            cAutoEject      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   720
         Index           =   0
         Left            =   1740
         Top             =   15
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "年 月"
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
         cPositionY      =   30
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   360
            Left            =   60
            TabIndex        =   19
            Top             =   315
            Width           =   1320
            _ExtentX        =   2328
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
            cSize           =   -1  'True
            Text            =   "200911"
            cChkResult      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   720
         Index           =   8
         Left            =   3180
         Top             =   15
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "　　　　　　　　取  引  先"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   300
            Left            =   105
            Top             =   15
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   529
            BackColor       =   14737632
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
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
            Begin VB.OptionButton H1Op_Torkb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "発注"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   990
               TabIndex        =   21
               Top             =   45
               Width           =   810
            End
            Begin VB.OptionButton H1Op_Torkb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "受注"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   105
               TabIndex        =   20
               Top             =   45
               Value           =   -1  'True
               Width           =   810
            End
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1080
            Top             =   330
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "テクノエイト"
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
            cAlingnment     =   1
            cPositionX      =   50
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            Top             =   315
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
            Text            =   "1046"
            MaxLength       =   7
         End
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   9990
      ScaleHeight     =   645
      ScaleWidth      =   5205
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1065
      Width           =   5265
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2610
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
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
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   3465
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
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
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   4320
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   45
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   900
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   5
         BtnCaption      =   "前進"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8421376
         BtnShadowColor  =   16777215
         PFCaption       =   "( Down )"
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
         PFBackColor     =   8421376
      End
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   1755
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   6
         BtnCaption      =   "後退"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8421376
         BtnShadowColor  =   16777215
         PFCaption       =   "( Up )"
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
         PFBackColor     =   8421376
      End
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 納入指示確認  】"
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
      cPositionX      =   1200
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14400
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   11805
         Top             =   45
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
         BackColor       =   14737632
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
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cBoderStyle     =   1
         cPositionX      =   50
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D7 
      Height          =   285
      Index           =   28
      Left            =   0
      Top             =   0
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      BackColor       =   12648447
      Caption         =   "220"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   12648447
      cBoderColor2    =   12648447
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cAlingnment     =   4
      cBoderStyle     =   2
   End
End
Attribute VB_Name = "CSA0070"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   納入指示確認･･･土岐津Ver
'**       フォームID    :   PSA0070
'**       処理概要      :
'**
'**       作  成  日    :   2010/01/22  By CIS
'**       変  更  日    :   2010/06/03  By CIS  取引先(入力項目)のみの場合ストアド追加
'**                                             グリッドの線引きを一括に変更
'**       変  更  日    :   2010/07/20  By CIS  変数ｵｰﾊﾞｰﾌﾛｰ修正
'**       変  更  日    :   2010/07/27  By CIS  ストアド⇒ワーク時、キー項目(受入)追加
'**       変  更  日    :   2010/08/23  By CIS  いろいろ変更
'**       変  更  日    :   2010/08/31  By CIS  引き落としを出荷→受注に変更
'**       変  更  日    :   2010/09/01  By CIS  いろいろ変更
'**       変  更  日    :   2010/11/11  By CIS  在庫メンテ移動時に品番セット
'**       変  更  日    :   2012/10/31  By CIS  ﾀｲﾑｱｳﾄ変更/受注ﾃﾞｰﾀ参照時の不具合修正
'**       変  更  日    :   2013/03/27  By CIS  在庫ﾏｽﾀ参照時の並び順変更
'**       変  更  日    :   2019/08/05  By CIS  在庫管理区分=3での在庫発注点＞在庫時の背景色変更対応
'**       変  更  日    :   2019/09/19  By CIS  グリッドの品番(材料)情報に背番号と得意先の表示を追加
'**       変  更  日    :
'**
'**********************************************************'
    Const c_MRow        As Integer = 5
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid1      As New CisVsGrid3
    Dim CisVsGrid2      As New CisVsGrid3

    Dim wDSP_YM         As String       '画面上の内示年月
    Dim wFLG_YM         As Integer      '画面上の対象年月..0/1/2
    Dim wKadoDay        As Currency     '稼動日数
    Dim WDd             As Integer

    Dim wDSP_JyuSu(50000, 72)      As Long         '1:出荷予定
    Dim wDSP_HacSu(50000, 72)      As Long         '2:入荷予定
    Dim wDSP_SyuSu(50000, 72)      As Long         '3:出荷実績
    Dim wDSP_NyuSu(50000, 72)      As Long         '4:入荷実績
    Dim wCMP_ZaiSu(50000, 72)      As Long         '5:在庫
    Dim wCMP_ZaiKb(50000, 72)      As String       '6:在庫区分
    Dim wCMP_NyuSu(50000, 72)      As Long         '7:入荷数／在庫計算用
    Dim wCMP_SyuSu(50000, 72)      As Long         '8:出荷数／在庫計算用
    
    Dim wDSP_Jyu_Nai(50000, 72)    As Byte         '9:内示数出力ﾌﾗｸﾞ
    Dim wDSP_Hac_Nai(50000, 72)    As Byte         '10:発注内示数出力ﾌﾗｸﾞ
    
    Dim wX              As Integer
    Dim wY              As Integer
    Dim wKensu          As Integer
    Dim wIndex          As Integer
    Dim wGIndex         As Integer
    Dim wHinbn          As String
    Dim wTorcd          As String
    Dim wZKanriNo       As Long
    Dim wZaisitu        As String
    Dim wItaatu         As Currency
    Dim wHaba           As Currency
    Dim wNagasa         As Currency

    Dim wDaySu(50000)      As Currency     ' 日当数
    Dim wLot(50000)        As Currency     ' ロット
    Dim wNaiSu0(50000)     As Currency     ' 当月内示数
    Dim wNaiSu1(50000)     As Currency     ' 翌月内示数
    Dim wNaiSu2(50000)     As Currency     ' 翌々月内示数
    
    Dim wZenNaiSu       As Currency     ' 前月内示数(大親)
    Dim wZenSyuSu       As Currency     ' 前月出荷数(大親)
    Dim wSyuSu          As Currency     ' 当月出荷数(大親)
    
    Dim wTimeSu(50000)     As Currency     ' 時間当数
    Dim wSyuyo(50000)      As Currency     ' 収容数
    Dim cZSu1(50000)       As Currency     ' 翌月　初在庫数
    Dim cZSu2(50000)       As Currency     ' 翌々月初在庫数
    
    Dim wHakosyu(50000)    As String       ' 荷姿
    Dim wOkiba(50000)      As String       ' 在庫場所
    Dim wSTani(50000)      As String       '>使用単位

    Dim cHinbn(50000)      As String       '
    Dim cZKanriNo(50000)   As Long         '
    Dim cTorcd(50000)      As String       '>取引先
    Dim cUkeir(50000)      As String       '>受入
    Dim cKoujyun(50000)    As Byte         '>工順

    Dim wYmdS           As String       '>開始日付（抽出用）
    Dim wYmdE           As String       '>終了日付（抽出用）
    Dim wDateS          As String       '>開始日付（表示用）
    Dim wDateE          As String       '>終了日付（標示用）
    Dim wDays           As Integer      '>表示期間：日
    Dim wYM0            As String       ' 当月
    Dim wYM1            As String       ' 翌月
    Dim wYM2            As String       ' 翌々月
'
    Dim wApl_TTL(15, 9) As String       '>ﾀｲﾄﾙ
    Dim wApl_EXE(15, 9) As String       '>ﾌﾟﾛｸﾞﾗﾑ名
    Dim wApl_PAM(15, 9) As String       '>ﾊﾟﾗﾒｰﾀ
    Dim wApl_CTL(15, 9) As String       '>表示制御
    Dim wApl_RowNo      As Long
    Dim wApl_ColNo      As Long
'
    Dim wHinABC         As String       ' 品番表示限定
    Dim wHinABC_Table   As String
    Dim wHinABC_Where   As String

    Dim wSyoriDspKB     As Byte         ' 表示内容 ( 0:取引先単位[受注] 1:取引先単位[発注] 2:取引先,品番単位[受注][構成/親品番]
'                                                    3:取引先,品番単位[受注][詳細] 4:取引先,品番単位[発注][構成/親品番] 5:取引先,品番単位[発注][詳細]
'                                                    6:取引先,品番単位[受注][構成/指定品番以下] 7:取引先,品番単位[発注][構成/指定品番以下]
    Dim wOyaHinbn       As String       ' 選択親品番
    Dim wOyaTorcd       As String       ' 親品番取引先(発注先)
    
    Dim wZenKadosu      As Currency     '前月稼働日数
    Dim wNKadosu        As Currency     '当月稼働日数
    Dim wN1Kadosu       As Currency     '翌月稼働日数
    Dim wN2Kadosu       As Currency     '翌々月稼働日数
    Dim wZenYM          As String       '前月
    Dim wN1YM           As String       '翌月
    Dim wN2YM           As String       '翌々月
    Dim wN1DaySu        As Long         '翌月日当数
    Dim wN2DaySu        As Long         '翌々月日当数
    Dim wN1KadoKB       As String       '翌月稼動区分
    Dim wN2KadoKB       As String       '翌々月稼動区分
    Dim wN0KadoKB       As String       '当月稼動区分
    Dim wN_1KadoKB       As String      '前月稼動区分
    
    Dim wStr            As String
    Dim wStr1           As String
    Dim wInt            As Integer
    Dim wInt1           As Integer
    
    Dim Key_Hinban      As String       ' 2010/06/03 Insert
    Dim Key_Torcd       As String       ' 2010/06/03 Insert
    Dim Key_Ukeir       As String       ' 2010/07/27 Insert
    
    Dim wDSP_Flg        As Byte         ' 2010/08/23 取引先単位､品番出力用
    Dim wNYMD           As String       ' 2010/09/01 親最大納入日
    Dim wKoseiSu        As Integer      ' 2010/09/01 構成数
    

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
'    If Not CisFun.Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    gConnectCount = 6
    
    CisDB.TimeOut = 3600                    '2012/10/31
    
    If Not IniGet Then End
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl

    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
    Dummy.Left = -1000
    Dummy.Enabled = False

'   #------------------#
'   #  ｸﾞﾘｯﾄﾞ初期処理   #
'   #------------------#
    Set CisVsGrid1 = New CisVsGrid3
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .SelectCol = 5
        .InitGet = False
        .Init
    End With
    Call GridInit1                          ' ｸﾞﾘｯﾄﾞ初期表示
'   #------------------#
'   #  ｸﾞﾘｯﾄﾞ初期処理   #
'   #------------------#
    Set CisVsGrid2 = New CisVsGrid3
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
       '.PatanMax = 4
        .SelectCol = 62
        .InitGet = False
        .Init
    End With
    Call GridInit2                          ' ｸﾞﾘｯﾄﾞ初期表示

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    iH1_YM = Format(Now, "YYYYMM")
    H1Op_Torkb(0).Value = True
    H1Op_Zaiko(0).Value = True
    
    wSyoriDspKB = 0
    H3_Area1.Visible = False
    Sel_OyaHinbn.Visible = False
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

'   >> 名称マスタ確認
    Call MeisyoInit
    Call ReadMenu

''   >> 品番表示限定のｾｯﾄ .. set where from parameter
'    Select Case RTrim$(Command)
'        Case "A", "B", "C"
'            wHinABC = RTrim$(Command)
'            wHinABC_Table = " left outer join 品番マスタ hnm on hnm.品番 = htm.品番"
'            Select Case wHinABC
'                Case "A": wHinABC_Where = " and isnull(hnm.ABC,'') in ('A')"
'                Case "B": wHinABC_Where = " and isnull(hnm.ABC,'') in ('A','B')"
'                Case "C": wHinABC_Where = " and isnull(hnm.ABC,'') in ('A','B','C')"
'            End Select
'        Case Else
'            wHinABC = ""
'            wHinABC_Table = ""
'            wHinABC_Where = ""
'    End Select
'    K_Sykbnm.Caption = wHinABC

    FormAct = Not FormAct
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
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    If ProcHB = "B1" Then
        If Push_PFKey(Key_Code, Shift) Then Key_Acc = True: Exit Function
    End If
    Select Case Key_Code
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                 'Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                 'Dummy.Enabled = False
                               End If
           Case Else:          Exit Function
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
    Call CisDB.DBDISConnect
'+---------------------+
'+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
'+---------------------+
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(1)
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call Grid_ESC
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        iH1_Hinbn.SetFocus
        wOyaHinbn = ""
        Sel_OyaHinbn.Visible = False: DoEvents
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "H2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        iH1_Hinbn.SetFocus
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "H3" Then
        ProcHB = "H2"
        Call Grid_ESC
        Call DispChange(ProcHB)
        H2Op_DspKB(0).SetFocus
        wOyaHinbn = ""
        Sel_OyaHinbn.Visible = False: DoEvents
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'*********************************
Private Sub Grid_ESC()
    If ProcHB = "B1" Then vsGrid2.SetFocus
    SendKeys "{HOME}"
    DoEvents
    With vsGrid2
        .Redraw = False
        For gInt = 0 To 51
             .Row = 0
             .Col = gInt + 2
             .CellBackColor = &HC0C0C0
             .TextMatrix(0, .Col) = "          " & Format(gInt + 1, "00")
        Next gInt
        .Redraw = True
    End With
    vsGrid1.Rows = 1
    vsGrid2.Rows = 1
    vsGrid1.Rows = 30
    vsGrid2.Rows = 50000
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "select htm.*"
        gSL_Select = gSL_Select & "  from 品番取引先マスタ htm"
        gSL_Select = gSL_Select & wHinABC_Table
        gSL_Select = gSL_Select & " where htm.取引先 + space(7 - len(htm.取引先)) + htm.品番 < "
        gSL_Select = gSL_Select & "'" & iH1_Torcd & Space(7 - CisFun.Len2(iH1_Torcd)) & iH1_Hinbn & "'"
        If H1Op_Torkb(0) Then
           gSL_Select = gSL_Select & "   and htm.品目 = 0"
        Else
           gSL_Select = gSL_Select & "   and htm.品目 = 1"
        End If
        gSL_Select = gSL_Select & wHinABC_Where
        gSL_Select = gSL_Select & " order by htm.取引先 desc"
        gSL_Select = gSL_Select & "        , htm.品番 desc"
        If HTMRead(gSL_Select, 1) Then
           iH1_Torcd = HTM.取引先
           iH1_Hinbn = HTM.品番
        Else
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     先頭レコードです            "
                .MB_Title = "【後退】"
                .MB_Button = OK
                .MBOX
            End With
        End If
    End If
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "select htm.*"
        gSL_Select = gSL_Select & "  from 品番取引先マスタ htm"
        gSL_Select = gSL_Select & wHinABC_Table
        gSL_Select = gSL_Select & " where htm.取引先 + space(7 - len(htm.取引先)) + htm.品番 > "
        gSL_Select = gSL_Select & "'" & iH1_Torcd & Space(7 - CisFun.Len2(iH1_Torcd)) & iH1_Hinbn & "'"
        If H1Op_Torkb(0) Then
           gSL_Select = gSL_Select & "   and htm.品目 = 0"
        Else
           gSL_Select = gSL_Select & "   and htm.品目 = 1"
        End If
        gSL_Select = gSL_Select & wHinABC_Where
        gSL_Select = gSL_Select & " order by htm.取引先"
        gSL_Select = gSL_Select & "        , htm.品番"
        If HTMRead(gSL_Select, 1) Then
           iH1_Torcd = HTM.取引先
           iH1_Hinbn = HTM.品番
        Else
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     最終レコードです            "
                .MB_Title = "【前進】"
                .MB_Button = OK
                .MBOX
            End With
        End If
    End If
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH2_Hinbn" Then Call Look_Hinbn: Exit Sub     '2010/08/23
    If PB_Look.Tag = "iH2_Tokui" Then Call Look_Tokui: Exit Sub     '2010/08/23
End Sub
''+----------------------------+
''+    年月切替
''+----------------------------+
'Private Sub B1Op_YM_Click(INDEX As Integer)
'    If ProcHB = "B1" Then
'        Call CisFrm.MousePT(11)
'        ProcHB = "B2"
'        PNL_PF.Enabled = False
'        B1_Area2.Enabled = False
'        B1_Area3.Enabled = False
'        Proc_MSG.Visible = True: DoEvents       '>ﾒｯｾｰｼﾞ表示
'        Call Grid_ESC
'        If INDEX = 0 Then
'            wDSP_YM = RTrim(iH1_YM)
'        Else
'            CisFun.Date_RsFormat = fYM
'            CisFun.Date_Add RTrim(iH1_YM), M, INDEX
'            wDSP_YM = Mid(CisFun.Date_Result, 1, 6)
'        End If
'        wFLG_YM = INDEX
'        Call Naiji_Kosei_Crt                    ' 表示データ作成
'        Call BodySet("")
'        Call SetPos_Grid2                       '>日別明細の表示開始ｾｯﾄ
'        Proc_MSG.Visible = False: DoEvents      '>ﾒｯｾｰｼﾞ消去
'        ProcHB = "B1"
'        PNL_PF.Enabled = True
'        B1_Area2.Enabled = True
'        B1_Area3.Enabled = True
''       >> メニュー内容の表示処理を実行 .. 2004/06/30
'        Call vsGrid2_RowColChange
''       -----------------------------
'        Call CisFrm.MousePT(1)
'    End If
'End Sub
'+-----------------------------+
'+       取引先検索             +
'+-----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    If H1Op_Torkb(0) Then
       RV_TorKb = 3
    Else
       RV_TorKb = 9
    End If

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_Torcd
        H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0

    If H1Op_Torkb(0) Then
       RVI_Hinmoku = 0     '>受注
    Else
       RVI_Hinmoku = 1     '>発注
    End If
    RVI_SK = 9
       
    RVI_Torcd = RTrim(iH1_Torcd)
    RV_Call = ""

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
'====================== 2010/08/23 Start
'       iH1_Hinbn = RV_Hinbn
'       iH1_Torcd = RV_TorcdK
'       Call iH1_Torcd_LostFocus
       If ProcHB = "H1" Then                                '2010/11/11
'       If PB_Look.Tag = "iH1_Hinbn" Then                   '2010/11/11
          iH1_Hinbn = RV_Hinbn
          iH1_Torcd = RV_TorcdK
          Call iH1_Torcd_LostFocus
       Else
          iH2_Hinbn = RV_Hinbn
       End If
'====================== 2010/08/23 End
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
    If H1Op_Torkb(0) Then
       If TorNmGet(RTrim(iH1_Torcd), 3) Then
           H1lb_Tornm = TRM_RName
       End If
    Else
       If TorNmGet(RTrim(iH1_Torcd), 9) Then
           H1lb_Tornm = TRM_RName
       End If
    End If
End Sub
'+-----------------------------+
'+  iH1_Hinbn
'+-----------------------------+
Private Sub iH1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  H1Op_Torkb   取引先区分
'+-----------------------------+
Private Sub H1Op_Torkb_Click(Index As Integer)
    H1Op_Torkb(0).BackColor = &HE0E0E0
    H1Op_Torkb(1).BackColor = &HE0E0E0
    H1Op_Torkb(Index).BackColor = &HC0FFC0
    H1Op_Torkb(1).Tag = Format(Index, "0")
End Sub
'+-----------------------------+
'+  H2Op_DspKB   構成区分
'+-----------------------------+
Private Sub H2Op_DspKB_Click(Index As Integer)
    H2Op_DspKB(0).BackColor = &HE0E0E0
    H2Op_DspKB(1).BackColor = &HE0E0E0
    H2Op_DspKB(Index).BackColor = &HC0FFC0
    H2Op_DspKB(1).Tag = Format(Index, "0")
    If Index = 1 Then
       'Back_DspKosei.Visible = False
       H2Op_DspKosei(0).Value = True
       Back_DspKosei.Enabled = False
    Else
       'Back_DspKosei.Visible = True
       H2Op_DspKosei(0).Value = True
       Back_DspKosei.Enabled = True
    End If
End Sub
'+-----------------------------+
'+  H2Op_DspKosei 構成詳細
'+-----------------------------+
Private Sub H2Op_DspKosei_Click(Index As Integer)
    H2Op_DspKosei(0).BackColor = &HE0E0E0
    H2Op_DspKosei(1).BackColor = &HE0E0E0
    H2Op_DspKosei(Index).BackColor = &HC0FFC0
    H2Op_DspKosei(1).Tag = Format(Index, "0")
End Sub
'==================== 2010/08/23 Start
'+-----------------------------+
'+  H2Op_DataKB データ有無
'+-----------------------------+
Private Sub H2Op_DataKB_Click(Index As Integer)
    H2Op_DataKB(0).BackColor = &HE0E0E0
    H2Op_DataKB(1).BackColor = &HE0E0E0
    H2Op_DataKB(Index).BackColor = &HC0FFC0
    H2Op_DataKB(1).Tag = Format(Index, "0")
End Sub
'+-----------------------------+
'+  iH2_Hinbn
'+-----------------------------+
Private Sub iH2_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH2_Tokui
'+-----------------------------+
Private Sub iH2_Tokui_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Tokui_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H2lb_Tokui = ""
    If TorNmGet(iH2_Tokui, 0) Then
       H2lb_Tokui = TRM_RName
    End If
End Sub
'+----------------------------+
'+       代表得意先検索       +
'+----------------------------+
Private Sub Look_Tokui()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 0
    RV_Call = ""
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
        iH2_Tokui = RV_TorcdK
        H2lb_Tokui = RV_TorRName
    End If
End Sub
'==================== 2010/08/23 End
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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim H3Mode      As Boolean
    Dim H3Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long

    H1Mode = False:     H1Color = gPLostSel
    H2Mode = False:     H2Color = gPLostSel
    H3Mode = False:     H3Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel

' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
'
    Select Case Pro
        Case "H1":  H1Mode = True:  H1Color = gPGotSel: PB_ENT.BtnCaption = "入力"
        Case "H2":  H2Mode = True:  H2Color = gPGotSel
        Case "H3":  H3Mode = True:  H3Color = gPGotSel
        Case "B1":  B1Mode = True:  B1Color = gPGotSel: PB_ENT.BtnCaption = "最新"
    End Select

    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode:  H2_Area1.BackColor = H2Color
    H3_Area1.Enabled = H3Mode:  H3_Area1.BackColor = H3Color
    B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color
    'B1_Area2.Enabled = B1Mode:  B1_Area2.BackColor = B1Color
    B1_Area3.Enabled = B1Mode:  B1_Area3.BackColor = B1Color
    B1_Area3.Visible = B1Mode
    H2_Area1.Visible = H2Mode Or H3Mode Or B1Mode
    H3_Area1.Visible = H3Mode

    PB_Look.Visible = False
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = B1Mode Or H2Mode Or H3Mode
    PB_ENT.Visible = H1Mode Or B1Mode Or H2Mode Or H3Mode
    
    If ProcHB = "H2" Then
       '===================== 2010/08/23 Start
       If RTrim(iH1_Hinbn) <> "" Then
          Back_Kobetu.Visible = True
          Back_TorKB.Visible = False
       Else
          Back_Kobetu.Visible = False
          Back_TorKB.Visible = True
       End If
       H2Op_DataKB(0).Value = True
       '===================== 2010/08/23 End
       
       H2Op_DspKB(0).Value = True
    End If

    If ProcHB <> "H2" Then
       If RTrim(iH1_Hinbn) = "" Then
          H2_Area1.Visible = False
       End If
    End If
    
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
                .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj

    If ProcHB = "H1" Then
       H2Op_DspKB(0).Value = True: H2Op_DspKosei(0).Value = True
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
'========== 2010/08/23 Del Start
'        If RTrim(iH1_Hinbn) = "" Then
'           ProcHB = "B1"
'           If H1Op_Torkb(0) Then
'              wSyoriDspKB = 0
'           Else
'              wSyoriDspKB = 1
'           End If
'           Call DispChange(ProcHB)
'           Call SetPos_Grid2           '>日別明細の表示開始ｾｯﾄ
'           GoTo ReturnPress_Ed
'        Else
'========== 2010/08/23 Del End
           '詳細条件入力
           ProcHB = "H2"
           Call DispChange(ProcHB)
           GoTo ReturnPress_Ed
'        End If
    End If
    If ProcHB = "H2" Then
       If H1Op_Torkb(0) Then
'======================== 2010/08/23 Start
'          If H2Op_DspKB(1) Then
'             wSyoriDspKB = 3
'          Else
'             If H2Op_DspKosei(0) Then
'                wSyoriDspKB = 2
'             Else
'                wSyoriDspKB = 6
'             End If
'          End If
          If RTrim(iH1_Hinbn) <> "" Then
             If H2Op_DspKB(1) Then
                wSyoriDspKB = 3
             Else
                If H2Op_DspKosei(0) Then
                   wSyoriDspKB = 2
                Else
                   wSyoriDspKB = 6
                End If
             End If
          Else
             wSyoriDspKB = 0
          End If
'======================== 2010/08/23 End
       Else
'======================== 2010/08/23 Start
'          If H2Op_DspKB(1) Then
'             wSyoriDspKB = 5
'          Else
'             If H2Op_DspKosei(0) Then
'                wSyoriDspKB = 4
'             Else
'                wSyoriDspKB = 7
'             End If
'          End If
          If RTrim(iH1_Hinbn) <> "" Then
             If H2Op_DspKB(1) Then
                wSyoriDspKB = 5
             Else
                If H2Op_DspKosei(0) Then
                   wSyoriDspKB = 4
                Else
                   wSyoriDspKB = 7
                End If
             End If
          Else
             wSyoriDspKB = 1
          End If
'======================== 2010/08/23 End
       End If
    
       If wSyoriDspKB = 4 Then
           If GyakuTenkai(RTrim(iH1_Hinbn), 0) = 1 Then
              ProcHB = "H3"
              Call DispChange(ProcHB)
              Call GridSet1
              vsGrid1.SetFocus
              Proc_MSG.Visible = False: DoEvents
              GoTo ReturnPress_Ed
           Else
              If Not Head2Chk Then
                 Call DispChange(ProcHB)
                 GoTo ReturnPress_Ed
              End If
           End If
       Else
           If Not Head2Chk Then
              Call DispChange(ProcHB)
              GoTo ReturnPress_Ed
           End If
       End If
    
       
       ProcHB = "B1"
       Call DispChange(ProcHB)
       Call SetPos_Grid2           '>日別明細の表示開始ｾｯﾄ
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "H3" Then
        '
        If RTrim(wOyaHinbn) = "" Then Call vsGrid1_Select(1)
        If RTrim(wOyaHinbn) = "" Then GoTo ReturnPress_Ed
        
        If Not Head2Chk Then GoTo ReturnPress_Ed
        
        Sel_OyaHinbn.Visible = True: DoEvents
        
        ProcHB = "B1"
        Call DispChange(ProcHB)
        Call SetPos_Grid2           '>日別明細の表示開始ｾｯﾄ
        GoTo ReturnPress_Ed
    End If
'    If ProcHB = "B1" Then
'        Call B1Op_YM_Click(wFLG_YM)
'    End If
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

    If RTrim(iH1_YM) = "" Then
       CisFun.MB_Lines = 4
       CisFun.MB_MSG(2) = "    年月を入力して下さい            "
       CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If RTrim(iH1_Torcd) = "" Then
       CisFun.MB_Lines = 4
       CisFun.MB_MSG(2) = "    取引先を入力して下さい            "
       CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If H1Op_Torkb(0) Then
       If Not TorNmGet(RTrim(iH1_Torcd), 3) Then
          CisFun.MB_Lines = 4
          CisFun.MB_MSG(2) = "    納入先が未登録です            "
          CisFun.MBOX
          iH1_Torcd.SetFocus
          Exit Function
       Else
          H1lb_Tornm = RTrim(TRM_RName)        '>略称
       End If
    Else
       If Not TorNmGet(RTrim(iH1_Torcd), 9) Then
          CisFun.MB_Lines = 4
          CisFun.MB_MSG(2) = "    取引先が未登録です            "
          CisFun.MBOX
          iH1_Torcd.SetFocus
          Exit Function
       Else
          H1lb_Tornm = RTrim(TRM_RName)        '>略称
       End If
    End If
    
    If H1Op_Torkb(0) Then
       wSyoriDspKB = 0
    Else
       wSyoriDspKB = 1
    End If

'    If RTrim(iH1_Hinbn) = "" Then
'       CisFun.MB_Lines = 4
'       CisFun.MB_MSG(2) = "    品番を入力して下さい            "
'       CisFun.MBOX
'       iH1_Torcd.SetFocus
'       Exit Function
'    End If
'   >> 品番取引先マスタ
    If RTrim(iH1_Hinbn) <> "" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "select *"
        gSL_Select = gSL_Select & "  from 品番取引先マスタ"
        gSL_Select = gSL_Select & " where 品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(iH1_Torcd) & "'"
        If H1Op_Torkb(0) Then
           gSL_Select = gSL_Select & "   and 品目 = 0"
        Else
           gSL_Select = gSL_Select & "   and 品目 = 1"
        End If
        If Not HTMRead(gSL_Select, 1) Then
           CisFun.MB_Lines = 4
           CisFun.MB_MSG(2) = "    取引先・品番　未登録！（品番取引先マスタ）            "
           CisFun.MBOX
           iH1_Torcd.SetFocus
           Exit Function
        End If
'        gSL_Select = ""
'        gSL_Select = gSL_Select & "select hnm.ABC"
'        gSL_Select = gSL_Select & "  from 品番マスタ hnm"
'        gSL_Select = gSL_Select & " where hnm.品番 = '" & RTrim(iH1_Hinbn) & "'"
'        gSL_Select = gSL_Select & wHinABC_Where
'        If Not HNMRead(gSL_Select, 1) Then
'           CisFun.MB_Lines = 4
'           CisFun.MB_MSG(2) = "    入力された品番は対象外です            "
'           CisFun.MBOX
'           Exit Function
'        End If
    End If

    gSL_Select = "SELECT * FROM カレンダマスタ "
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1 "
    gSL_Select = gSL_Select & " AND 取引先 = ''"
    gSL_Select = gSL_Select & " AND 年 = '" & Mid(iH1_YM, 1, 4) & "'"
    gSL_Select = gSL_Select & " AND 月 = '" & Mid(iH1_YM, 5, 2) & "'"
    If CLMRead(gSL_Select, 1) Then
       wKadoDay = CLM.稼動日数
    Else
       CisFun.MB_Lines = 4
       CisFun.MB_MSG(2) = "    カレンダマスタ未登録【自社】            "
       CisFun.MBOX
       iH1_YM.SetFocus
       Exit Function
    End If
    
'   >> 年月のｾｯﾄ
    With CisFun
        wYM0 = iH1_YM                        '>当月
        .Date_RsFormat = fYM
        .Date_Add wYM0, M, 1
        wN1YM = Mid(.Date_Result, 1, 6)      '>翌月
        .Date_Add wYM0, M, 2
        wN2YM = Mid(.Date_Result, 1, 6)      '>翌々月
        .Date_Add wYM0, M, -1
        wZenYM = Mid(.Date_Result, 1, 6)    '>前月
    End With


' マスタ入力 SELECT文 ｾｯﾄ
    '=========== 取引先のみ入力
    If RTrim(iH1_Hinbn) = "" Then
        wDSP_YM = RTrim(iH1_YM)
'        ProcHB = "H2"
        Proc_MSG.Visible = True: DoEvents
    
        wFLG_YM = 0
'        Call Naiji_Kosei_Crt0        '表示データ作成
        
        ' テスト用                                                                  2010/06/02 Insert
'        If wSyoriDspKB = 0 Or wSyoriDspKB = 1 Then
'            Call NouData_Edit
'        Else
'            Call Naiji_Kosei_Crt0        '表示データ作成
'        End If
        

'=================- 2010/08/23 Del Start
'        gSL_Select = "SELECT * FROM 納入指示構成ワーク "
'        gSL_Select = gSL_Select & " WHERE 処理端末 = SubString(Host_Name(),1,20)"
'        If Not DBInput("INV") Then
'            GoTo Head1Chk_ED
'        End If
    
        'B1Op_YM(0).Value = True

    '   <><><><><> 月初在庫数の退避 <><><><><>
'        For gLong = 0 To 99
'            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wYM0 & "01", "@@@@/@@/@@"))
'            cZSu1(gLong) = wCMP_ZaiSu(gLong, WDd)
'            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wYM1 & "01", "@@@@/@@/@@"))
'            cZSu2(gLong) = wCMP_ZaiSu(gLong, WDd)
'        Next gLong
'=================- 2010/08/23 Del End
    
    End If
'   -------------------------------------
    Head1Chk = True
Head1Chk_ED:
    Proc_MSG.Visible = False: DoEvents
    ProcHB = "H1"
End Function
'****************************
'*      HEAD2ﾁｪｯｸ処理
'****************************
Private Function Head2Chk() As Boolean
    Head2Chk = False

'    If RTrim(????) = "" Then
'       CisFun.MB_Lines = 4
'       CisFun.MB_MSG(2) = "    年月を入力して下さい            "
'       CisFun.MBOX
'       iH1_Torcd.SetFocus
'       Exit Function
'    End If

    If RTrim(iH1_Hinbn) = "" Then GoTo Head2Chk_10          '2010/08/23

' マスタ入力 SELECT文 ｾｯﾄ
    wDSP_YM = RTrim(iH1_YM)
    If Proc_MSG.Visible = False Then
       Proc_MSG.Visible = True: DoEvents
    End If

    wFLG_YM = 0
    If wSyoriDspKB = 3 Or wSyoriDspKB = 5 Then
       Call Naiji_Kosei_Crt1        '表示データ作成(詳細)
    Else
       If wSyoriDspKB = 2 Or wSyoriDspKB = 6 Or wSyoriDspKB = 7 Then
          wOyaHinbn = RTrim(iH1_Hinbn)
       End If
       
       Call Naiji_Kosei_Crt2        '表示データ作成(構成情報)
    End If
    
'================================ 2010/08/23 Start
'    gSL_Select = "SELECT * FROM 納入指示構成ワーク "
'    gSL_Select = gSL_Select & " WHERE 処理端末 = SubString(Host_Name(),1,20)"
'    If Not DBInput("INV") Then
'        GoTo Head2Chk_ED
'    End If
'
''   <><><><><> 月初在庫数の退避 <><><><><>
'    For gLong = 0 To 99
'        WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wYM0 & "01", "@@@@/@@/@@"))
'        cZSu1(gLong) = wCMP_ZaiSu(gLong, WDd)
'        WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wYM1 & "01", "@@@@/@@/@@"))
'        cZSu2(gLong) = wCMP_ZaiSu(gLong, WDd)
'    Next gLong
'    Head2Chk = True
    
    GoTo Head2Chk_20

Head2Chk_10:
    wDSP_YM = RTrim(iH1_YM)
    Proc_MSG.Visible = True: DoEvents

    wFLG_YM = 0
    Call Naiji_Kosei_Crt0        '表示データ作成
    

Head2Chk_20:
    gSL_Select = "SELECT * FROM 納入指示構成ワーク "
    gSL_Select = gSL_Select & " WHERE 処理端末 = SubString(Host_Name(),1,20)"
    If Not DBInput("INV") Then
        GoTo Head2Chk_ED
    End If

'   <><><><><> 月初在庫数の退避 <><><><><>
    For gLong = 0 To 99
        WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wYM0 & "01", "@@@@/@@/@@"))
        cZSu1(gLong) = wCMP_ZaiSu(gLong, WDd)
        WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wYM1 & "01", "@@@@/@@/@@"))
        cZSu2(gLong) = wCMP_ZaiSu(gLong, WDd)
    Next gLong
    Head2Chk = True

'================================ 2010/08/23 End

Head2Chk_ED:
    Proc_MSG.Visible = False: DoEvents
    ProcHB = "H1"
End Function
'****************************
'*      HEAD3ﾁｪｯｸ処理
'****************************
Private Function Head3Chk() As Boolean

    Head3Chk = False
    
    Head3Chk = True
    
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    If Not NSKRead(gSL_Select, 1) Then
       If InputKey = "Back" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "     先頭レコードです            "
               .MB_Title = "【後退】"
               .MB_Button = OK
               .MBOX
           End With
           GoTo DBInput_Ed
       End If
       If InputKey = "Next" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "     最終レコードです            "
               .MB_Title = "【前進】"
               .MB_Button = OK
               .MBOX
           End With
           GoTo DBInput_Ed
       End If
       If InputKey = "DSP" Or InputKey = "INV" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "     データが存在しません          "
               .MB_Title = ""
               .MB_Button = OK
               .MBOX
           End With
           GoTo DBInput_Ed
       End If
    Else
        Call BodySet("")
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)

    'Call GridSet1
'   >> 年月のｾｯﾄ
    With CisFun
        wYM0 = wDSP_YM                      '>当月
        .Date_RsFormat = fYM
        .Date_Add wYM0, M, 1
        wYM1 = Mid(.Date_Result, 1, 6)      '>翌月
        .Date_Add wYM0, M, 2
        wYM2 = Mid(.Date_Result, 1, 6)      '>翌々月
        .Date_Add wYM0, M, -1
        wZenYM = Mid(.Date_Result, 1, 6)    '>前月
    End With


    Call WeekSet

    Call GridSet2

    Call Grid2_Day_Set

End Sub
'+--------------------------------+
'+      日別明細の表示開始ｾｯﾄ
'+--------------------------------+
Private Sub SetPos_Grid2()
'   >> 行位置ｾｯﾄ
    vsGrid2.Row = 1
'   >> 日付ｶﾗﾑ位置のｾｯﾄ
    If wDSP_YM = Format(Now(), "yyyymm") Then
       vsGrid2.Col = DateDiff("d", Format(wDateS, "@@@@/@@/@@"), Now) + 4
    Else
       vsGrid2.Col = 17
    End If
'    If vsGrid2.Rows < 30 Then vsGrid2.Col = vsGrid2.Col + 1
    vsGrid2.Col = vsGrid2.Col + 1
    DoEvents
'   >> 日付ｶﾗﾑの移動
    SendKeys "{RIGHT}"
    SendKeys "{LEFT}"
    SendKeys "{LEFT}"
    DoEvents
'    If vsGrid2.Rows < 30 Then SendKeys "{LEFT}"
    SendKeys "{LEFT}"
    vsGrid2.Refresh
'   当月以外
    If wDSP_YM <> Format(Now(), "yyyymm") Then
        SendKeys "{LEFT}"
        SendKeys "{LEFT}"
        SendKeys "{LEFT}"
    End If
End Sub

'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
'       .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 左中, あり, 3, 0, " ")
        Call .FixedSet(中中, 左中, あり, 40, 1, "親品番")
        Call .FixedSet(中中, 左中, あり, 25, 2, "親品番")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 3, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
        .InitDisp
    End With
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
'       .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        
'< 2006.01.19 MNT START >-------------------------------------------------------------*
'        Call .FixedSet(中中, 左中, あり, 1, 0, "")
        Call .FixedSet(中中, 右中, あり, 2, 0, "")
'< 2006.01.19 MNT  END  >-------------------------------------------------------------*
        
'        Call .FixedSet(中中, 左中, あり, 25, 1, "品番（材質）")
        Call .FixedSet(中中, 左中, あり, 21, 1, "品番（材質）")         '2010/07/23
'        Call .FixedSet(中中, 中中, あり, 4, 2, "日程")
        Call .FixedSet(中中, 中中, あり, 8, 2, "日程")                  '2010/07/23

        gStr = "0,1,2"
        For gInt = 3 To 53
            Call .FixedSet(中中, 右中, あり, 7, gInt, "          " & Format(gInt - 2, "00"))
            gStr = gStr & "," & gInt
        Next gInt

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 中中, なし, 0, 54, "X品番")
        Call .FixedSet(中中, 左中, なし, 0, 55, "X仕入先")
        Call .FixedSet(中中, 左中, なし, 0, 56, "X部品材料区分")
        Call .FixedSet(中中, 左中, なし, 0, 57, "X材質")
        Call .FixedSet(中中, 左中, なし, 0, 58, "X板厚")
        Call .FixedSet(中中, 左中, なし, 0, 59, "X幅")
        Call .FixedSet(中中, 左中, なし, 0, 60, "X長さ")
        Call .FixedSet(中中, 左中, なし, 0, 61, "X社内区分")
        Call .FixedSet(中中, 左中, なし, 0, 62, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, gStr)
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
        .InitDisp
    End With
End Sub
Private Sub GridSet1(Optional DBNo As Integer = 1)
    Dim wIndex      As Integer
    
    gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示品番 FROM 構成問合せワーク KW "
    gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
    gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示品番"
    gSL_Select = gSL_Select & " ORDER BY KW.親品番"
    
    If Not KOLRead(gSL_Select, , 1) Then
        Call KOLClose(DBNo)
        Exit Sub
    Else
        wIndex = 1
        Do Until Not KOL_RDSTS
            With vsGrid1
                .Rows = CisDB.RecordCount + 1
                .TextMatrix(wIndex, 0) = Format(wIndex, "#0")
                .TextMatrix(wIndex, 1) = KOL.表示品番
                .TextMatrix(wIndex, 2) = KOL.親品番
            End With
        
            Call KOLReadNext(DBNo)
            wIndex = wIndex + 1
        Loop
        
        Call KOLClose(DBNo)
    End If

End Sub
Private Sub vsGrid1_DblClick()
    
    Call vsGrid1_Select(0)
    
End Sub
Private Sub vsGrid1_Select(KB As Byte)
    wOyaHinbn = ""
    With vsGrid1
        If .TextMatrix(.Row, 1) <> "" Then
            wOyaHinbn = RTrim(.TextMatrix(.Row, 2))
            Sel_OyaHinbn = "親品番：" & RTrim(.TextMatrix(.Row, 1))
        End If
    End With
    
    If KB = 0 Then Call ReturnPress
    
End Sub
''******************************************'
''*****       グリッド内容セット        *****
''******************************************'
'Private Sub GridSet1()
'    Dim ID          As Long
'    Dim GridCol()   As Long
'    Dim Ttl_Kensu   As Long
'    Dim sDoubl      As Double
'
'' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
'    With vsGrid1
'       .Redraw = False
'        ReDim GridCol(.Cols)
'        For ID = 0 To .Cols - 1
'            GridCol(ID) = CisVsGrid1.FixedGet(ID)
'        Next ID
'    End With
''   >> 背景色のｾｯﾄ
'    For gCnt1 = 2 To 2
'        vsGrid1.Row = gCnt1
'        For gCnt2 = 1 To 4
'            vsGrid1.Col = gCnt2
'            vsGrid1.CellBackColor = &HC0FFC0
'        Next gCnt2
'    Next gCnt1
'
'    Erase wDaySu
'    Erase wLot
'    Erase wNaiSu0
'    Erase wNaiSu1
'    Erase wNaiSu2
'    Erase wTimeSu
'    Erase wSyuyo
'    Erase wHakosyu
'    Erase wOkiba
'    Erase wSTani        '>使用単位
'
''   >> 年月のｾｯﾄ
'    With CisFun
'        wYM0 = wDSP_YM                      '>当月
'        .Date_RsFormat = fYM
'        .Date_Add wYM0, M, 1
'        wYM1 = Mid(.Date_Result, 1, 6)      '>翌月
'        .Date_Add wYM0, M, 2
'        wYM2 = Mid(.Date_Result, 1, 6)      '>翌々月
'    End With
'
'    gSL_Select = ""
'
''-----------------------------------------------------------------------< 2004.11.10 MNT START >---------------------
'
''    gSL_Select = gSL_Select & "select"
''    gSL_Select = gSL_Select & "  nsk.*"
''    gSL_Select = gSL_Select & " from"
''    gSL_Select = gSL_Select & "  納入指示構成ワーク nsk"
'
'    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  nsk.*, Isnull( zr.取数, 0 ) 取数"
'    gSL_Select = gSL_Select & " from"
'    gSL_Select = gSL_Select & "  納入指示構成ワーク nsk"
'    gSL_Select = gSL_Select & " left outer join 材料品番マスタ zr"
'    gSL_Select = gSL_Select & " on  zr.材料管理番号 = nsk.材料管理番号"
'    gSL_Select = gSL_Select & " and zr.使用品番 = nsk.品番"
'
''-----------------------------------------------------------------------< 2004.11.10 MNT  END  >---------------------
'
'    gSL_Select = gSL_Select & " where nsk.処理端末 = substring(host_name( ), 1, 20)"
'    gSL_Select = gSL_Select & "   and nsk.展開区分 in (0, 1)"
'    gSL_Select = gSL_Select & " order by"
'    gSL_Select = gSL_Select & "  nsk.順位"
'    gSL_Select = gSL_Select & ", nsk.品番"
'
'    wIndex = 0: wGIndex = 1
'    With vsGrid1
'        If NSGRead(gSL_Select) Then
'
'        .Rows = (5 * CisDB.RecordCount) + 1
'        vsGrid2.Rows = .Rows
'
'            Do Until NSG_RDSTS = False
'                If wIndex = 0 Then
'                    Call Naiji_NJT_Get      '>完成品：内示テーブル
'                Else
'                    Select Case NSG.部品材料区分
'                        Case "B":   Call Naiji_HNJ_Get      '>部品：発注内示テーブル ........ 手配先単位の時: (RTrim(NSG.仕入先)) を指定
'                        Case "Z":   Call Naiji_ZNJ_Get      '>材料：材料発注内示テーブル
'                    End Select
'                End If
'
'                'VSGridに転送
'                '【見出し】
'                .Row = wGIndex
'                .Col = 1: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex, 1) = "ﾋﾜﾘ / ﾛｯﾄ"
'                .Col = 2: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex, 2) = "当月"
'                .Col = 3: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex, 3) = "翌月"
'                .Col = 4: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex, 4) = "翌々月"
'                .Row = wGIndex + 2
'                If NSG.部品材料区分 = "B" Then
'                    .Col = 1: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 1) = "時間当"
'                    .Col = 2: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 2) = "収容数"
'                Else
'
''< 2004.11.10 MNT START >---------------------------------------------------------------------------------------------
'
''                    .Col = 1: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 1) = "使用量"
''                    .Col = 2: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 2) = "使用単位"
'                    If NSG.長さ <> 0 Then
'                        .Col = 1: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 1) = "取数"
'                        .Col = 2: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 2) = "単位"
'                    Else
'                        .Col = 1: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 1) = "使用量"
'                        .Col = 2: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 2) = "使用単位"
'                    End If
'
''< 2004.11.10 MNT  END  >---------------------------------------------------------------------------------------------
'
'                End If
'                .Col = 3: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 3) = "荷姿"
'                .Col = 4: .CellBackColor = &HFFC0C0: .TextMatrix(wGIndex + 2, 4) = "在庫場所"
'                '【ＢＯＤＹ】
'                .TextMatrix(wGIndex + 1, 1) = Format(wDaySu(wIndex), "###") & "/" & Format(wLot(wIndex), "###")
'                If NSG.部品材料区分 = "B" Then
'                    .TextMatrix(wGIndex + 1, 2) = CisFun.RSetFld(wNaiSu0(wIndex), 10, "#####")
'                    .TextMatrix(wGIndex + 1, 3) = CisFun.RSetFld(wNaiSu1(wIndex), 10, "#####")
'                    .TextMatrix(wGIndex + 1, 4) = CisFun.RSetFld(wNaiSu2(wIndex), 10, "#####")
'                    .TextMatrix(wGIndex + 3, 1) = CisFun.RSetFld(wTimeSu(wIndex), 8, "#####")
'                    .TextMatrix(wGIndex + 3, 2) = CisFun.RSetFld(wSyuyo(wIndex), 8, "####")
'                Else
'                    If wNaiSu0(wIndex) <> 0 Then .TextMatrix(wGIndex + 1, 2) = CisFun.RSetFld(wNaiSu0(wIndex), 10, "##0.000")
'                    If wNaiSu1(wIndex) <> 0 Then .TextMatrix(wGIndex + 1, 3) = CisFun.RSetFld(wNaiSu1(wIndex), 10, "##0.000")
'                    If wNaiSu2(wIndex) <> 0 Then .TextMatrix(wGIndex + 1, 4) = CisFun.RSetFld(wNaiSu2(wIndex), 10, "##0.000")
'
''< 2004.11.10 MNT START >---------------------------------------------------------------------------------------------
'
''                    .TextMatrix(wGIndex + 3, 1) = CisFun.RSetFld(NSG.構成数, 9, "##0.0000")
''                    .TextMatrix(wGIndex + 3, 2) = RTrim(wSTani(wIndex))
'
'                    If NSG.長さ <> 0 Then
'                        .TextMatrix(wGIndex + 3, 1) = CisFun.RSetFld(NSG.取数, 9, "#,##0")
'                        .TextMatrix(wGIndex + 3, 2) = "個"
'                    Else
'                        .TextMatrix(wGIndex + 3, 1) = CisFun.RSetFld(NSG.構成数, 9, "##0.0000")
'                        .TextMatrix(wGIndex + 3, 2) = RTrim(wSTani(wIndex))
'                    End If
'
''< 2004.11.10 MNT  END  >---------------------------------------------------------------------------------------------
'
'                End If
'                .TextMatrix(wGIndex + 3, 3) = wHakosyu(wIndex)
'                .TextMatrix(wGIndex + 3, 4) = wOkiba(wIndex)
'            '   >> 備考の編集
'                For gInt = 1 To 4
'                    If RTrim(NSG.備考) = "" Then
'                        .TextMatrix(wGIndex + 4, gInt) = " "
'                    Else
'                        .TextMatrix(wGIndex + 4, gInt) = NSG.備考 & Space(39 - CisFun.Len2(NSG.備考))
'                    End If
'                Next gInt
'                .MergeRow(wGIndex + 4) = True
'
'                .Row = wGIndex + 4
'                For gInt = 1 To 4
'                    .Col = gInt
'                    .CellBorder &H0&, 0, 0, 0, 2, 0, 0
'                Next gInt
'
'                If wIndex = 0 Then
''                   If iH1_YM = Format(Now, "YYYYMM") Then
''                      B1lb_YoteiSu = Format(wDaySu(wIndex) * CalKado(iH1_YM, "", "", 0), "###,###,##0")
''                   Else
''                      B1lb_YoteiSu = Format(wDaySu(wIndex) * CalKadoLast(iH1_YM, "", "", 0), "###,###,##0")
''                   End If
''                   B1lb_SyukaSu = Format(wSyuSu, "###,###,##0")
'
''                   If CisFun.Val2(B1lb_SyukaSu) = 0 Or CisFun.Val2(B1lb_YoteiSu) = 0 Then
''                      sDoubl = 0
''                   Else
''                      sDoubl = (CisFun.Val2(B1lb_SyukaSu) / CisFun.Val2(B1lb_YoteiSu)) * 100
''                      sDoubl = CisFun.HasuProc(sDoubl, 四捨五入, 2)
''                   End If
''                   B1lb_Hiritu = Format(sDoubl, "##0.00")
''                   If sDoubl < 100 Then
''                      B1lb_Hiritu.ForeColor = &HFF&                 '赤
''                   Else
''                      B1lb_Hiritu.ForeColor = &HFF0000              '青
''                   End If
'                End If
'
'                wIndex = wIndex + 1
'                wGIndex = wGIndex + 5
'                Call NSGReadNext
'
'            Loop
'        End If
'        Call NSGClose
'        .MergeCells = flexMergeRestrictRows
'        .Redraw = True
'
'        .Rows = wGIndex
'    End With
'
'End Sub
Private Function CalKado(cYM As String, cTorcd As String, cUke As String, CKb As Byte) As Integer
    Dim sI      As Integer
    CalKado = 0
        
    gSL_Select = "Select * From カレンダマスタ"
    gSL_Select = gSL_Select & " Where 取引先区分 = " & CKb & ""
    gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(cTorcd) & "'"
    gSL_Select = gSL_Select & "   And 年 = '" & Left(cYM, 4) & "'"
    gSL_Select = gSL_Select & "   And 月 = '" & Mid(cYM, 5, 2) & "'"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(cUke) & "'"
    If Not CLMRead(gSL_Select, 1, 1) Then
       Exit Function
    Else
        For sI = 1 To CInt(Format(Now, "DD"))
            If Mid(CLM.稼動区分, sI, 1) = "0" Then
               CalKado = CalKado + 1
            End If
        Next sI
    End If

End Function
Private Function CalKadoLast(cYM As String, cTorcd As String, cUke As String, CKb As Byte) As Integer
    Dim sI      As Integer
    CalKadoLast = 0
        
    gSL_Select = "Select * From カレンダマスタ"
    gSL_Select = gSL_Select & " Where 取引先区分 = " & CKb & ""
    gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(cTorcd) & "'"
    gSL_Select = gSL_Select & "   And 年 = '" & Left(cYM, 4) & "'"
    gSL_Select = gSL_Select & "   And 月 = '" & Mid(cYM, 5, 2) & "'"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(cUke) & "'"
    If Not CLMRead(gSL_Select, 1, 1) Then
       Exit Function
    Else
        For sI = 1 To 31
            If Mid(CLM.稼動区分, sI, 1) = "0" Then
               CalKadoLast = CalKadoLast + 1
            End If
        Next sI
    End If

End Function
'+-------------------------------------------------------------
'+      内示テーブル - 内示情報
'+  変更    2005/10/05  by TOTEC    当月内示月計数を日別内示数の集計に変更
'+          2006/07/25  By CIS      前月内示数取得
'+-------------------------------------------------------------
Private Sub Naiji_NJT_Get(Optional DBNo As Integer = 1)
    Dim iDay    As Integer
    Dim sZYM    As String
'   >> 内示数
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  njt.内示数"
    gSL_Select = gSL_Select & ", njt.翌月内示数"
    gSL_Select = gSL_Select & ", njt.翌々月内示数"
    '↓--------------------------当月日別内示数の集計用に取得------------
    For iDay = 1 To 31
        gSL_Select = gSL_Select & ", njt.内示数量" & Format(iDay, "00")
    Next iDay
    '↑--------------------------当月日別内示数の集計用に取得------------
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  内示テーブル njt"
    gSL_Select = gSL_Select & " where njt.内示年月 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   and njt.品番 = '" & RTrim(NSG.大親品番) & "'"
    gSL_Select = gSL_Select & "   and njt.納入先 = '" & RTrim(NSG.大親納入先) & "'"
    If NJTRead(gSL_Select, 1, DBNo) Then
        Select Case wFLG_YM
            Case 0
                '↓----------------------------------当月内示月計数を日別内示数の集計に変更----
                'wNaiSu0(wIndex) = NJT.内示数
                wNaiSu0(wIndex) = 0
                For iDay = 1 To 31
                    wNaiSu0(wIndex) = wNaiSu0(wIndex) + NJT.内示数量(iDay)
                Next iDay
                '↑----------------------------------当月内示月計数を日別内示数の集計に変更----
                wNaiSu1(wIndex) = NJT.翌月内示数
                wNaiSu2(wIndex) = NJT.翌々月内示数
            Case 1
                wNaiSu0(wIndex) = NJT.翌月内示数
                wNaiSu1(wIndex) = NJT.翌々月内示数
                wNaiSu2(wIndex) = 0
            Case 2
                wNaiSu0(wIndex) = NJT.翌々月内示数
                wNaiSu1(wIndex) = 0
                wNaiSu2(wIndex) = 0
        End Select
    Else
        wNaiSu0(wIndex) = 0
        wNaiSu1(wIndex) = 0
        wNaiSu2(wIndex) = 0
    End If
    
    With CisFun
        .Date_Add RTrim(iH1_YM) & "01", M, -1
        sZYM = Mid(.Date_Result, 1, 6)
    End With
    
    If wFLG_YM = 0 Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  njt.内示数"
        gSL_Select = gSL_Select & ", njt.翌月内示数"
        gSL_Select = gSL_Select & ", njt.翌々月内示数"
        For iDay = 1 To 31
            gSL_Select = gSL_Select & ", njt.内示数量" & Format(iDay, "00")
        Next iDay
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  内示テーブル njt"
        gSL_Select = gSL_Select & " where njt.内示年月 = '" & RTrim(sZYM) & "'"
        gSL_Select = gSL_Select & "   and njt.品番 = '" & RTrim(NSG.大親品番) & "'"
        gSL_Select = gSL_Select & "   and njt.納入先 = '" & RTrim(NSG.大親納入先) & "'"
        
        If NJTRead(gSL_Select, 1, DBNo) Then
           wZenNaiSu = 0
           For iDay = 1 To 31
               wZenNaiSu = wZenNaiSu + NJT.内示数量(iDay)
           Next iDay
        Else
            wZenNaiSu = 0
        End If
    End If
    
'   >> 日当数算出
    wDaySu(wIndex) = wNaiSu0(wIndex) / wKadoDay

'   >> 部品情報の取得（時間当＆収容数＆荷姿＆ロット＆在庫場所＆備考）
'    Call Buhin_Info_Get(DBNo)

'   >> ロットはクリア
    wLot(wIndex) = 0

End Sub
'+-------------------------------------------------------------
'+      発注内示テーブル - 内示情報
'+-------------------------------------------------------------
Private Sub Naiji_HNJ_Get(Optional Torcd As String = "", Optional DBNo As Integer = 1)
'   >> 内示数合計算出
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  sum(case when substring(hnj.納期, 1, 6) = '" & wYM0 & "' then hnj.数量 else 0 end) 当月内示数"
    gSL_Select = gSL_Select & ", sum(case when substring(hnj.納期, 1, 6) = '" & wYM1 & "' then hnj.数量 else 0 end) 翌月内示数"
    gSL_Select = gSL_Select & ", sum(case when substring(hnj.納期, 1, 6) = '" & wYM2 & "' then hnj.数量 else 0 end) 翌々月内示数"
    gSL_Select = gSL_Select & " from"
    
'    gSL_Select = gSL_Select & "  発注内示テーブル hnj"
    gSL_Select = gSL_Select & "  発注内示必要数テーブル hnj"
'
    gSL_Select = gSL_Select & " where hnj.内示年月 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   and hnj.品番 = '" & RTrim(NSG.品番) & "'"
    If Torcd <> "" Then
        gSL_Select = gSL_Select & "   and hnj.手配先 = '" & Torcd & "'"
    End If
    Call HNJRead(gSL_Select, 1, DBNo)
    wNaiSu0(wIndex) = HNJ.当月内示数
    wNaiSu1(wIndex) = HNJ.翌月内示数
    wNaiSu2(wIndex) = HNJ.翌々月内示数
'   >> 日当数算出
    wDaySu(wIndex) = wNaiSu0(wIndex) / wKadoDay

'   >> 部品情報の取得（時間当＆収容数＆荷姿＆ロット＆在庫場所＆備考）
'    Call Buhin_Info_Get(DBNo)

End Sub
''+-------------------------------------------------------------
''+      部品情報（時間当・収容数・荷姿・ロット・在庫場所・備考）
''+-------------------------------------------------------------
'Private Sub Buhin_Info_Get(Optional DBNo As Integer = 1)
''   >> 時間当数獲得
'    gSL_Select = "SELECT * FROM T_品番取引作業拡張マスタ "
'    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(NSG.品番) & "'"
'    gSL_Select = gSL_Select & " AND 取引先 = '" & RTrim(NSG.仕入先) & "'"
'    gSL_Select = gSL_Select & " AND 作業順 = ( SELECT MAX(作業順) FROM 品番工順マスタ "
'    gSL_Select = gSL_Select & "                              WHERE 品番 = '" & RTrim(NSG.品番) & "'"
'    gSL_Select = gSL_Select & "                              AND 取引先 = '" & RTrim(NSG.仕入先) & "' )"
'    If HTKRead(gSL_Select, 1, DBNo) Then
'        wTimeSu(wIndex) = HTK.時間当生産数
'    Else
'        wTimeSu(wIndex) = 0
'    End If
''   >> 収容数＆荷姿＆ロット＆在庫場所
'    gSL_Select = "SELECT * FROM 品番取引先マスタ "
'    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(NSG.品番) & "'"
'    gSL_Select = gSL_Select & " AND 取引先 = '" & RTrim(NSG.仕入先) & "'"
'    gSL_Select = gSL_Select & " AND 品目 = 1 "
'    gSL_Select = gSL_Select & " ORDER BY 受入"
'    If HTMRead(gSL_Select, 1, DBNo) Then
'        If RTrim(HTM.変更日) <> "" And _
'           RTrim(HTM.変更日) <= wDSP_YM & "31" Then
'            wSyuyo(wIndex) = HTM.収容数2
'            wHakosyu(wIndex) = HTM.収容器2
'            wLot(wIndex) = HTM.ロット2
'        Else
'            wSyuyo(wIndex) = HTM.収容数1
'            wHakosyu(wIndex) = HTM.収容器1
'            wLot(wIndex) = HTM.ロット1
'        End If
'    '   >> 在庫場所
'        wOkiba(wIndex) = RTrim(HTM.置場)
'        cTorcd(wIndex) = RTrim(HTM.取引先)
'        cUkeir(wIndex) = RTrim(HTM.受入)
'    '   >> 備考
'        NSG.備考 = HTM.備考
'    Else
'        wLot(wIndex) = 0
'        wSyuyo(wIndex) = 0
'        wHakosyu(wIndex) = ""
'        wOkiba(wIndex) = ""
'        cTorcd(wIndex) = ""
'        cUkeir(wIndex) = ""
'    End If
'End Sub
'+-------------------------------------------------------------
'+      材料発注内示テーブル - 内示情報
'+-------------------------------------------------------------
Private Sub Naiji_ZNJ_Get(Optional DBNo As Integer = 1)
'   >> 内示数合計算出
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  sum(case when substring(znj.納期, 1, 6) = '" & wYM0 & "' then znj.数量 else 0 end) 当月内示数"
    gSL_Select = gSL_Select & ", sum(case when substring(znj.納期, 1, 6) = '" & wYM1 & "' then znj.数量 else 0 end) 翌月内示数"
    gSL_Select = gSL_Select & ", sum(case when substring(znj.納期, 1, 6) = '" & wYM2 & "' then znj.数量 else 0 end) 翌々月内示数"
'
'    gSL_Select = gSL_Select & " from 材料発注内示テーブル znj"
    gSL_Select = gSL_Select & " from 材料発注内示必要数テーブル znj"
'
    gSL_Select = gSL_Select & " where znj.内示年月 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   and znj.材料管理番号 = " & NSG.材料管理番号 & ""
    Call ZNJRead(gSL_Select, 1, DBNo)
    wNaiSu0(wIndex) = ZNJ.当月内示数
    wNaiSu1(wIndex) = ZNJ.翌月内示数
    wNaiSu2(wIndex) = ZNJ.翌々月内示数
'   >> 日当数算出
    'wDaySu(wIndex) = 0                             '材料の日当所要量も稼働日の単純割りで算出
    wDaySu(wIndex) = wNaiSu0(wIndex) / wKadoDay     '材料の日当所要量も稼働日の単純割りで算出

'   >> ロット
'   >> 時間当数獲得
    wTimeSu(wIndex) = 0
'   >> 収容数
    wSyuyo(wIndex) = 0
'   >> 荷姿
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  isnull(sy1.値名称, '') 単位"
    gSL_Select = gSL_Select & ", isnull(sy2.値名称, '') 手配単位"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料マスタ zrm"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sy1"
    gSL_Select = gSL_Select & "    on sy1.区分名称 = '単位区分'"
    gSL_Select = gSL_Select & "   and sy1.区分ﾀｲﾌﾟ = ''"
    gSL_Select = gSL_Select & "   and sy1.値 = zrm.単位"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sy2"
    gSL_Select = gSL_Select & "    on sy2.区分名称 = '単位区分'"
    gSL_Select = gSL_Select & "   and sy2.区分ﾀｲﾌﾟ = ''"
    gSL_Select = gSL_Select & "   and sy2.値 = zrm.手配単位"
    gSL_Select = gSL_Select & " where zrm.材料管理番号 = " & NSG.材料管理番号
    If ZRMRead(gSL_Select, 1, DBNo) Then
       wHakosyu(wIndex) = RTrim(ZRM.手配単位)
       wSTani(wIndex) = RTrim(ZRM.単位)
    End If
'   >> 在庫場所
'   __________// 下記の記述はとりあえずコメントにして、置場に変更 //__________
   'gSL_Select = "SELECT MIN(取引先) 取引先 FROM 材料取引先マスタ "
   'gSL_Select = gSL_Select & " WHERE 材料管理番号 = " & NSG.材料管理番号
   'If ZTMRead(gSL_Select, 1, DBNo) Then
   '   Call TorNmGet(RTrim(ZTM.取引先), 7, 1)
   '   wOkiba(wIndex) = CisFun.Mid2(RTrim(TRM_Name), 1, 10)
   'Else
   '   wOkiba(wIndex) = ""
   'End If
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  ztm.置場"
    gSL_Select = gSL_Select & ", ztm.取引先"
    gSL_Select = gSL_Select & ", ztm.受入"
    gSL_Select = gSL_Select & ", ztm.備考"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料取引先マスタ ztm"
    gSL_Select = gSL_Select & " where ztm.材料管理番号 = " & NSG.材料管理番号 & ""
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  ztm.取引先"
    gSL_Select = gSL_Select & ", ztm.受入"
    If ZTMRead(gSL_Select, 1, DBNo) Then
       wOkiba(wIndex) = RTrim(ZTM.置場)
       cTorcd(wIndex) = RTrim(ZTM.取引先)
       cUkeir(wIndex) = RTrim(ZTM.受入)
    '   >> 備考
        NSG.備考 = ZTM.備考
    Else
       wOkiba(wIndex) = ""
       cTorcd(wIndex) = ""
       cUkeir(wIndex) = ""
    End If
End Sub
'******************************************'
'*      見出し項目の設定
'******************************************'
Private Sub WeekSet()
    Dim wb_SDay     As Byte
    Dim wb_EDay     As Byte
    Dim wb_Day      As Integer
    Dim wl_ColNo    As Long

    wl_ColNo = 2

    wZenKadosu = 0
    wNKadosu = 0
    wN1Kadosu = 0
    wN2Kadosu = 0
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "select clm.*"
    gSL_Select = gSL_Select & "  from カレンダマスタ clm"
    gSL_Select = gSL_Select & " where clm.取引先区分 = 1"
    gSL_Select = gSL_Select & "   and clm.取引先 = ''"
    gSL_Select = gSL_Select & "   and clm.年 + clm.月"
    gSL_Select = gSL_Select & "       between '" & CisFun.Mid2(wDateS, 1, 6) & "'"
    gSL_Select = gSL_Select & "           and '" & CisFun.Mid2(wDateE, 1, 6) & "'"
    gSL_Select = gSL_Select & " order by clm.年"
    gSL_Select = gSL_Select & "        , clm.月"
    If CLMRead(gSL_Select) Then
        vsGrid2.Redraw = False
        Do Until CLM_RDSTS = False
            '前月稼働日数
            If RTrim(CLM.年) & RTrim(CLM.月) = RTrim(wZenYM) Then
               wZenKadosu = CLM.稼動日数
            End If
            '当月稼働日数
            If RTrim(CLM.年) & RTrim(CLM.月) = RTrim(iH1_YM) Then
               wNKadosu = CLM.稼動日数
            End If
            '翌月稼働日数
            If RTrim(CLM.年) & RTrim(CLM.月) = RTrim(wYM1) Then
               wN1Kadosu = CLM.稼動日数
            End If
            '翌々月稼働日数
            If RTrim(CLM.年) & RTrim(CLM.月) = RTrim(wYM2) Then
               wN2Kadosu = CLM.稼動日数
            End If
        
        '   >> 処理月の判定
            If CLM.月 = CisFun.Mid2(wDateS, 5, 2) Then
            '   [前月]
                wb_SDay = CisFun.Mid2(wDateS, 7, 2)
                wb_EDay = CLM.末日
                
                wN_1KadoKB = CLM.稼動区分
            Else
                If CLM.月 = CisFun.Mid2(wDateE, 5, 2) Then
                '   [翌月]
                    wb_SDay = 1
                    wb_EDay = 10
                    
                    wN1KadoKB = CLM.稼動区分
                Else
                '   [当月]
                    wb_SDay = 1
                    wb_EDay = CLM.末日
                
                    wN0KadoKB = CLM.稼動区分
                End If
            End If
            For wb_Day = wb_SDay To wb_EDay
                With vsGrid2
                    .Row = 0
                    wl_ColNo = wl_ColNo + 1
                    .TextMatrix(0, wl_ColNo) = CLM.月 & "/" & Format(wb_Day, "00")
                    If CisFun.Mid2(CLM.稼動区分, wb_Day, 1) = "1" Then
                        .Col = wl_ColNo
                        .CellBackColor = &HC0C0FF
                    End If
                    .Row = 1
                End With
            Next wb_Day
            Call CLMReadNext
        Loop
        vsGrid2.Redraw = True
    End If
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet2()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    Dim Ttl_Hinbn   As String   '>品番（材質）の編集用
    Dim Ttl_Sunpo   As String   '>寸法（板厚・幅・長さ）の編集用

    Dim Ttl_Seban   As String   '>背番号の編集用        ' 2019/09/19 ADD
    Dim Ttl_Tokui   As String   '>代表得意先の編集用    ' 2019/09/19 ADD

'   >> 背景色＆ｶﾗﾑ長のｾｯﾄ
    With vsGrid2
        .Redraw = False
        If wSyoriDspKB >= 2 Then
           If wSyoriDspKB <> 3 And wSyoriDspKB <> 5 And wSyoriDspKB <> 7 Then
               For gCnt1 = 1 To 2
                   .Row = gCnt1
'                   For gCnt2 = 2 To 53
'                       .Col = gCnt2
'                       .CellBackColor = &HC0FFC0
'                   Next gCnt2
               Next gCnt1
           End If
        End If
        For gCnt2 = 51 To 53
            .ColWidth(gCnt2) = .ColWidth(3)
        Next gCnt2
        Select Case DateDiff("d", Format(wDateS, "@@@@/@@/@@"), Format(wDateE, "@@@@/@@/@@")) + 1
            Case 48: .ColWidth(51) = 0: .ColWidth(52) = 0: .ColWidth(53) = 0
            Case 49: .ColWidth(52) = 0: .ColWidth(53) = 0
            Case 50: .ColWidth(53) = 0
        End Select
    End With

' 抽出処理 初期ｾｯﾄ
    wIndex = 1
    
    gSL_Select = "SELECT NS.*,ISNULL(TR.略称,'') 仕入先名,ISNULL(SY1.値名称,'') 作業名"
    
    '=======================================================================================' 2019/09/19 ADD START
    gSL_Select = gSL_Select + vbCrLf & "      ,CASE WHEN NS.部品材料区分 = 'B'"
    gSL_Select = gSL_Select + vbCrLf & "            THEN ISNULL(HT.背番号, '')"
    gSL_Select = gSL_Select + vbCrLf & "            ELSE ISNULL(ZT.背番号, '')"
    gSL_Select = gSL_Select + vbCrLf & "       END  背番号"
    gSL_Select = gSL_Select + vbCrLf & "      ,CASE WHEN NS.部品材料区分 = 'B'"
    gSL_Select = gSL_Select + vbCrLf & "            THEN ISNULL(HM.得意先, '')"
    gSL_Select = gSL_Select + vbCrLf & "            ELSE ISNULL(ZT.得意先, '')"
    gSL_Select = gSL_Select + vbCrLf & "       END  得意先"
    gSL_Select = gSL_Select + vbCrLf & "      ,CASE WHEN NS.部品材料区分 = 'B'"
    gSL_Select = gSL_Select + vbCrLf & "            THEN ISNULL(TB.略称  , '')"
    gSL_Select = gSL_Select + vbCrLf & "            ELSE ISNULL(TZ.略称  , '')"
    gSL_Select = gSL_Select + vbCrLf & "       END  得意先名"
    '=======================================================================================' 2019/09/19 ADD END

    gSL_Select = gSL_Select + vbCrLf & ", ISNULL(SY2.値名称,'') 形状名"
    gSL_Select = gSL_Select + vbCrLf & ", ISNULL(ZM.寸法区分, 0) 寸法区分"
    gSL_Select = gSL_Select + vbCrLf & ", ISNULL(TR.社内区分, 0) 社内区分"
    gSL_Select = gSL_Select + vbCrLf & ", ISNULL(sy3.値名称,'') 機械名"
    gSL_Select = gSL_Select + vbCrLf & ", ISNULL(sy4.値名称,'') 機械名称"
    gSL_Select = gSL_Select + vbCrLf & " FROM 納入指示構成ワーク NS"
    gSL_Select = gSL_Select + vbCrLf & " LEFT OUTER JOIN 取引先マスタ TR "
    gSL_Select = gSL_Select + vbCrLf & "   ON TR.取引先CD = NS.仕入先 "
    gSL_Select = gSL_Select + vbCrLf & "  AND TR.取引先区分 = 1 "
    gSL_Select = gSL_Select + vbCrLf & " LEFT OUTER JOIN 材料マスタ ZM "
    gSL_Select = gSL_Select + vbCrLf & "   ON ZM.材料管理番号 = NS.材料管理番号 "
    gSL_Select = gSL_Select + vbCrLf & " LEFT OUTER JOIN 名称マスタ SY1 "
    gSL_Select = gSL_Select + vbCrLf & "   ON SY1.区分名称 = '作業管理'"
    gSL_Select = gSL_Select + vbCrLf & "  AND SY1.値 = NS.加工CD "
    gSL_Select = gSL_Select + vbCrLf & " LEFT OUTER JOIN 名称マスタ SY2 "
    gSL_Select = gSL_Select + vbCrLf & "   ON SY2.区分名称 = '形状区分'"
    gSL_Select = gSL_Select + vbCrLf & "  AND SY2.値 = ZM.形状 "
    gSL_Select = gSL_Select + vbCrLf & " left outer join 品番工順マスタ htk"
    gSL_Select = gSL_Select + vbCrLf & "   on htk.品番 = NS.品番"
    gSL_Select = gSL_Select + vbCrLf & "  and htk.取引先 = NS.仕入先"
    gSL_Select = gSL_Select + vbCrLf & "  and htk.受入 = NS.受入"
    gSL_Select = gSL_Select + vbCrLf & "  and htk.工順 = NS.作業順"
    gSL_Select = gSL_Select + vbCrLf & " left outer join 名称マスタ sy3"
    gSL_Select = gSL_Select + vbCrLf & "   on sy3.区分名称 = '機械NO'"
    gSL_Select = gSL_Select + vbCrLf & "  and sy3.値 = htk.機械NO"
    gSL_Select = gSL_Select + vbCrLf & " left outer join 名称マスタ sy4"
    gSL_Select = gSL_Select + vbCrLf & "   on sy4.区分名称 = '機械'"
    gSL_Select = gSL_Select + vbCrLf & "  and sy4.値 = htk.機械"

    '=======================================================================================' 2019/09/19 ADD START
    gSL_Select = gSL_Select + vbCrLf & "/* 部品の背番号と得意先の取得 */"
    gSL_Select = gSL_Select + vbCrLf & "  LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select + vbCrLf & "    ON HT.品番   = NS.品番"
    gSL_Select = gSL_Select + vbCrLf & "   AND HT.取引先 = NS.大親納入先"
    gSL_Select = gSL_Select + vbCrLf & "  LEFT OUTER JOIN 品番マスタ HM"
    gSL_Select = gSL_Select + vbCrLf & "    ON HM.品番   = NS.品番"
    gSL_Select = gSL_Select + vbCrLf & "  LEFT OUTER JOIN 取引先マスタ TB"
    gSL_Select = gSL_Select + vbCrLf & "    ON TB.取引先CD   = HM.得意先"
    gSL_Select = gSL_Select + vbCrLf & "   AND TB.取引先区分 = 0"
    gSL_Select = gSL_Select + vbCrLf & "/* 材料の背番号と得意先の取得 */"
  ''gSL_Select = gSL_Select + vbCrLf & "  LEFT OUTER JOIN 材料取引先マスタ ZT"                                      ' 2019/09/20 DEL
    gSL_Select = gSL_Select + vbCrLf & "  LEFT OUTER JOIN"                                                          ' 2019/09/20 ADD
    gSL_Select = gSL_Select + vbCrLf & "                 (SELECT TOP(1) 背番号, 材料管理番号, 取引先, 得意先"       ' 2019/09/20 ADD
    gSL_Select = gSL_Select + vbCrLf & "                    FROM 材料取引先マスタ"                                  ' 2019/09/20 ADD
    gSL_Select = gSL_Select + vbCrLf & "                   GROUP BY 背番号, 材料管理番号, 取引先, 得意先"           ' 2019/09/20 ADD
    gSL_Select = gSL_Select + vbCrLf & "                   ORDER BY 背番号 DESC) ZT"                                ' 2019/09/20 ADD
    gSL_Select = gSL_Select + vbCrLf & "    ON ZT.材料管理番号 = NS.材料管理番号"
    gSL_Select = gSL_Select + vbCrLf & "   AND ZT.取引先 = NS.大親納入先"
    gSL_Select = gSL_Select + vbCrLf & "  LEFT OUTER JOIN 取引先マスタ TZ"
    gSL_Select = gSL_Select + vbCrLf & "    ON TZ.取引先CD   = ZT.得意先"
    gSL_Select = gSL_Select + vbCrLf & "   AND TZ.取引先区分 = 0"
    '=======================================================================================' 2019/09/19 ADD END

    gSL_Select = gSL_Select + vbCrLf & " WHERE NS.処理端末 = SubString(Host_Name(),1,20)"
    gSL_Select = gSL_Select + vbCrLf & "   AND NS.展開区分 in (0, 1)"
    If wSyoriDspKB = 3 Or wSyoriDspKB = 5 Then
       gSL_Select = gSL_Select + vbCrLf & " ORDER BY NS.作業順, NS.品番, NS.部品材料区分"
    Else
       gSL_Select = gSL_Select + vbCrLf & " ORDER BY NS.順位, NS.品番, NS.部品材料区分"
    End If

    Call Str_Monitor(gSL_Select)
    If NSGRead(gSL_Select) Then
        Do Until NSG_RDSTS = False
            With vsGrid2
'                .TextMatrix(wIndex, 2) = "出予"                    '2010/07/23
'                .TextMatrix(wIndex + 1, 2) = "出庫"
'                .TextMatrix(wIndex + 2, 2) = "入予"
'                .TextMatrix(wIndex + 3, 2) = "入庫"
                
                .TextMatrix(wIndex, 2) = "受注数"
                .TextMatrix(wIndex + 1, 2) = "出荷数"
                .TextMatrix(wIndex + 2, 2) = "入庫予定"
                .TextMatrix(wIndex + 3, 2) = "生産実績"
                .TextMatrix(wIndex + 4, 2) = "現在在庫"

                If NSG.部品材料区分 = "B" Then
                    If wSyoriDspKB >= 2 Then
                        '===============================================================================================================' 2019/09/19 DEL START
'                       Ttl_Hinbn = RTrim(NSG.品番) & vbCrLf _
'                              & vbCrLf _
'                              & CisFun.Mid2(RTrim(NSG.仕入先名), 1, 12) & "　" & CisFun.Mid2(RTrim(NSG.作業名), 1, 12) & vbCrLf _
'                              & vbCrLf _
'                              & CisFun.Mid2(RTrim(NSG.機械名称), 1, 20) & "   " & CisFun.Mid2(RTrim(NSG.機械名), 1, 20)
                        '===============================================================================================================' 2019/09/19 DEL END
                        '===============================================================================================================' 2019/09/19 ADD START
                        Ttl_Seban = ""
                        If Trim(NSG.背番号) <> "" Then
                            Ttl_Seban = "背番号：" & NSG.背番号
                        End If
                        Ttl_Tokui = ""
                        If Trim(NSG.得意先名) <> "" Then
                            Ttl_Tokui = "得意先：" & NSG.得意先名
                        End If
                        Ttl_Hinbn = RTrim(NSG.品番) & vbCrLf _
                                  & CisFun.Mid2(RTrim(NSG.仕入先名), 1, 12) & "　" & CisFun.Mid2(RTrim(NSG.作業名), 1, 12) & vbCrLf _
                                  & CisFun.Mid2(RTrim(NSG.機械名称), 1, 20) & "   " & CisFun.Mid2(RTrim(NSG.機械名), 1, 20) & vbCrLf _
                                  & vbCrLf _
                                  & Ttl_Seban & vbCrLf _
                                  & CisFun.Mid2(Ttl_Tokui, 1, 21)
                        '===============================================================================================================' 2019/09/19 ADD END
                    Else
                       Ttl_Hinbn = RTrim(NSG.品番)
                    End If
                Else
                    If NSG.寸法区分 = 0 Then
                        If NSG.長さ = 0 Then
                            Ttl_Sunpo = Format(NSG.板厚, "###0.00") & " × " & Format(NSG.幅, "###0.00")
                        Else
                            Ttl_Sunpo = Format(NSG.板厚, "###0.00") & " × " & Format(NSG.幅, "###0.00") & " × " & Format(NSG.長さ, "###0.00")
                        End If
                    Else
                        Ttl_Sunpo = ""
                    End If
                    '===================================================================' 2019/09/19 DEL START
'                    Ttl_Hinbn = RTrim(NSG.形状名) & vbCrLf _
'                              & RTrim(NSG.材質) & vbCrLf _
'                              & Ttl_Sunpo & vbCrLf _
'                              & vbCrLf _
'                              & "           " & RTrim(NSG.使用品番)
                    '===================================================================' 2019/09/19 DEL END
                    '===================================================================' 2019/09/19 ADD START
                    Ttl_Seban = ""
                    If Trim(NSG.背番号) <> "" Then
                        Ttl_Seban = "背番号：" & NSG.背番号
                    End If
                    Ttl_Tokui = ""
                    If Trim(NSG.得意先名) <> "" Then
                        Ttl_Tokui = "得意先：" & NSG.得意先名
                    End If
                    Ttl_Hinbn = RTrim(NSG.形状名) & vbCrLf _
                              & RTrim(NSG.材質) & vbCrLf _
                              & Ttl_Sunpo & vbCrLf _
                              & "           " & RTrim(NSG.使用品番) & vbCrLf _
                              & Ttl_Seban & vbCrLf _
                              & CisFun.Mid2(Ttl_Tokui, 1, 21)
                    '===================================================================' 2019/09/19 ADD END
                End If
            '   品番情報の退避
                For gInt = 0 To 4
                    If wSyoriDspKB >= 2 Then
                       If wSyoriDspKB <> 3 And wSyoriDspKB <> 5 Then
                            .TextMatrix(wIndex + gInt, 0) = vbCrLf & vbCrLf & vbCrLf _
                                                           & NSG.レベル & vbCrLf _
                                                           & vbCrLf & vbCrLf & vbCrLf _
                                                           & Ttl_Hinbn
                       Else
                            .TextMatrix(wIndex + gInt, 0) = vbCrLf & vbCrLf & vbCrLf _
                                                           & NSG.作業順 & vbCrLf _
                                                           & vbCrLf & vbCrLf & vbCrLf _
                                                           & Ttl_Hinbn
                       End If
                    Else
                       .TextMatrix(wIndex + gInt, 0) = vbCrLf & vbCrLf & vbCrLf _
                                                      & " " & vbCrLf _
                                                      & vbCrLf & vbCrLf & vbCrLf _
                                                      & Ttl_Hinbn
                    End If
                    .TextMatrix(wIndex + gInt, 1) = Ttl_Hinbn
                    .TextMatrix(wIndex + gInt, 54) = RTrim(NSG.品番)
                    .TextMatrix(wIndex + gInt, 55) = RTrim(NSG.仕入先)
                    .TextMatrix(wIndex + gInt, 56) = RTrim(NSG.部品材料区分)
                    .TextMatrix(wIndex + gInt, 57) = RTrim(NSG.材質)
                    .TextMatrix(wIndex + gInt, 58) = NSG.板厚
                    .TextMatrix(wIndex + gInt, 59) = NSG.幅
                    .TextMatrix(wIndex + gInt, 60) = NSG.長さ
                    .TextMatrix(wIndex + gInt, 61) = NSG.社内区分
                Next gInt

                .Row = wIndex + 4
'--------------------------------------------- 2010/06/03 Update Start
'                For gInt = 2 To 53
'                    .Col = gInt
'                    .CellBorder &H0&, 0, 0, 0, 2, 0, 0
'                Next gInt
                .Col = 2
                .ColSel = 53
                .CellBorder &H0&, 0, 0, 0, 2, 0, 0
'--------------------------------------------- 2010/06/03 Update End
            End With

            wIndex = wIndex + 5
            Call NSGReadNext

        Loop

        vsGrid2.MergeCol(1) = True
        For gInt = 2 To 53
            vsGrid2.MergeCol(gInt) = False
        Next gInt
        vsGrid2.MergeCells = flexMergeRestrictColumns

        vsGrid2.Row = 1: vsGrid2.Col = 1
        vsGrid2.Redraw = True

        vsGrid2.Rows = wIndex
        'vsGrid2.Rows = wIndex + 5

    End If
    Call NSGClose

End Sub
'**************************************************************
'*      日別明細セット
'**************************************************************
Private Sub Grid2_Day_Set()
    Dim wb_ZaiKB    As Boolean
    Dim wl_Zaiko    As Long
    Dim wl_ColNo    As Long
    Dim wi_End      As Integer

    wIndex = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wDateS, "@@@@/@@/@@"))
    wi_End = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wYmdE, "@@@@/@@/@@")) + 1

    vsGrid2.Redraw = False

'    For wY = 0 To wKensu
    For wY = 0 To wKensu - 1
'        If wY = 33 Then
'           Debug.Print wY
'        End If
        wb_ZaiKB = False
        Select Case vsGrid2.TextMatrix((wY * 5) + 1, 56)
            Case "B": wb_ZaiKB = Zaiko_Set("B")
            Case "Z": wb_ZaiKB = Zaiko_Set("Z")
        End Select
        wl_Zaiko = wCMP_ZaiSu(wY, 0)
        Debug.Print "元在庫数 = " & Format(wl_Zaiko, "###,##0")
        
        For wX = 1 To wi_End
'            If wCMP_NyuSu(wY, wX) = 600 Then
'               MsgBox "aaa"
'            End If
        '   >> 在庫計算
            If wb_ZaiKB Then
                Select Case wCMP_ZaiKb(wY, wX)
'                    Case "0":   wl_Zaiko = wCMP_ZaiSu(wY, wX) + wCMP_NyuSu(wY, wX) - wCMP_SyuSu(wY, wX)
'                    Case "1":   wl_Zaiko = wCMP_ZaiSu(wY, wX)
'                    Case Else:  wl_Zaiko = wl_Zaiko + wCMP_NyuSu(wY, wX) - wCMP_SyuSu(wY, wX)
                
                    Case "0":
                                If wDSP_Jyu_Nai(wY, wX) = 1 Then
                                    If wDSP_Hac_Nai(wY, wX) = 1 Then
'                                       wl_Zaiko = wCMP_ZaiSu(wY, wX) - wCMP_SyuSu(wY, wX)                              '2010/08/31
                                       wl_Zaiko = wCMP_ZaiSu(wY, wX) - wDSP_JyuSu(wY, wX)                               '2010/08/31
                                    Else
'                                       wl_Zaiko = wCMP_ZaiSu(wY, wX) + wCMP_NyuSu(wY, wX) - wCMP_SyuSu(wY, wX)         '2010/08/31
                                       wl_Zaiko = wCMP_ZaiSu(wY, wX) + wCMP_NyuSu(wY, wX) - wDSP_JyuSu(wY, wX)          '2010/08/31
                                    End If
                                Else
                                    If wDSP_Hac_Nai(wY, wX) = 1 Then
'                                       wl_Zaiko = wCMP_ZaiSu(wY, wX) - wCMP_SyuSu(wY, wX)                              '2010/08/31
'                                       wl_Zaiko = wCMP_ZaiSu(wY, wX) - wDSP_JyuSu(wY, wX)                      '2010/09/01                         '2010/08/31
                                       wl_Zaiko = wCMP_ZaiSu(wY, wX) + wCMP_NyuSu(wY, wX) - wDSP_JyuSu(wY, wX)  '2010/09/01          '2010/08/31
                                    Else
'                                       wl_Zaiko = wCMP_ZaiSu(wY, wX) + wCMP_NyuSu(wY, wX) - wCMP_SyuSu(wY, wX)         '2010/08/31
                                       wl_Zaiko = wCMP_ZaiSu(wY, wX) + wCMP_NyuSu(wY, wX) - wDSP_JyuSu(wY, wX)          '2010/08/31
                                    End If
                                End If
                    
                    Case "1":   wl_Zaiko = wCMP_ZaiSu(wY, wX)
                    Case Else:
                                If wDSP_Jyu_Nai(wY, wX) = 1 Then
                                    If wDSP_Hac_Nai(wY, wX) = 1 Then
'                                       wl_Zaiko = wl_Zaiko - wCMP_SyuSu(wY, wX)                                        '2010/08/31
                                       wl_Zaiko = wl_Zaiko - wDSP_JyuSu(wY, wX)                                         '2010/08/31
                                    Else
'                                       wl_Zaiko = wl_Zaiko + wCMP_NyuSu(wY, wX) - wCMP_SyuSu(wY, wX)                   '2010/08/31
                                       wl_Zaiko = wl_Zaiko + wCMP_NyuSu(wY, wX) - wDSP_JyuSu(wY, wX)                    '2010/08/31
                                    End If
                                Else
                                    If wDSP_Hac_Nai(wY, wX) = 1 Then
'                                       wl_Zaiko = wl_Zaiko - wCMP_SyuSu(wY, wX)                                        '2010/08/31
'                                       wl_Zaiko = wl_Zaiko - wDSP_JyuSu(wY, wX)                        '2010/09/01                               '2010/08/31
                                       wl_Zaiko = wl_Zaiko + wCMP_NyuSu(wY, wX) - wDSP_JyuSu(wY, wX)    '2010/09/01                '2010/08/31
                                    Else
'                                       wl_Zaiko = wl_Zaiko + wCMP_NyuSu(wY, wX) - wCMP_SyuSu(wY, wX)                   '2010/08/31
                                       wl_Zaiko = wl_Zaiko + wCMP_NyuSu(wY, wX) - wDSP_JyuSu(wY, wX)                    '2010/08/31
                                    End If
                                End If
                
                End Select
                wCMP_ZaiSu(wY, wX) = wl_Zaiko
'                Debug.Print "カウント = " & Format(wX, "##") & "   入荷数 = " & Format(wCMP_NyuSu(wY, wX), "###,##0") & "  出荷数 = " & Format(wDSP_JyuSu(wY, wX), "###,##0") & "   在庫数 = " & Format(wl_Zaiko, "###,##0")
'                Debug.Print "カウント = " & Format(wX, "##") & "   入荷数 = " & Format(wCMP_NyuSu(wY, wX), "###,##0") & "  受注数 = " & Format(wDSP_JyuSu(wY, wX), "###,##0") & "   在庫数 = " & Format(wl_Zaiko, "###,##0")
            End If
        '   >> 表示判定
            If wX > wIndex Then
                wl_ColNo = wX - wIndex + 2
                With vsGrid2
'                    .TextMatrix((wY * 5) + 1, wl_ColNo) = Format(wDSP_JyuSu(wY, wX), "#####")
'                    .TextMatrix((wY * 5) + 2, wl_ColNo) = Format(wDSP_SyuSu(wY, wX), "#####")
'                    .TextMatrix((wY * 5) + 3, wl_ColNo) = Format(wDSP_HacSu(wY, wX), "#####")
'                    .TextMatrix((wY * 5) + 4, wl_ColNo) = Format(wDSP_NyuSu(wY, wX), "#####")
                    If H1Op_Zaiko(0) Then
                        If wDSP_Jyu_Nai(wY, wX) = 1 Then
                            If wDSP_JyuSu(wY, wX) <> 0 Then
                               .TextMatrix((wY * 5) + 1, wl_ColNo) = "*" & Format(wDSP_JyuSu(wY, wX), "#####")
                            Else
                               .TextMatrix((wY * 5) + 1, wl_ColNo) = "*" & Format(wDSP_SyuSu(wY, wX), "#####")
                            End If
                            .TextMatrix((wY * 5) + 2, wl_ColNo) = "*" & Format(wDSP_SyuSu(wY, wX), "#####")
                        Else
                            .TextMatrix((wY * 5) + 1, wl_ColNo) = Format(wDSP_JyuSu(wY, wX), "#####")
                            .TextMatrix((wY * 5) + 2, wl_ColNo) = Format(wDSP_SyuSu(wY, wX), "#####")
                        End If
                        
                        If wDSP_Hac_Nai(wY, wX) = 1 Then
'未来在庫表示に使用しない  2010/07/23
'                            .TextMatrix((wY * 5) + 3, wl_ColNo) = "*" & Format(wDSP_HacSu(wY, wX), "#####")
'                            .TextMatrix((wY * 5) + 4, wl_ColNo) = "*" & Format(wDSP_NyuSu(wY, wX), "#####")
                            .TextMatrix((wY * 5) + 3, wl_ColNo) = ""
                            .TextMatrix((wY * 5) + 4, wl_ColNo) = ""
                        Else
                            .TextMatrix((wY * 5) + 3, wl_ColNo) = Format(wDSP_HacSu(wY, wX), "#####")
                            .TextMatrix((wY * 5) + 4, wl_ColNo) = Format(wDSP_NyuSu(wY, wX), "#####")
                        End If
                     Else
                        .TextMatrix((wY * 5) + 1, wl_ColNo) = Format(wDSP_JyuSu(wY, wX), "#####")
                        .TextMatrix((wY * 5) + 2, wl_ColNo) = Format(wDSP_SyuSu(wY, wX), "#####")
                        .TextMatrix((wY * 5) + 3, wl_ColNo) = Format(wDSP_HacSu(wY, wX), "#####")
                        .TextMatrix((wY * 5) + 4, wl_ColNo) = Format(wDSP_NyuSu(wY, wX), "#####")
                     End If
                                            
'                    wInt = .Row
'                    If wDSP_Jyu_Nai(wY, wX) = 1 Then
'                       .Col = wl_ColNo: .Row = (wY * 5) + 1
'                       .CellBackColor = &HFFC0FF
'                       .Col = wl_ColNo: .Row = (wY * 5) + 2
'                       .CellBackColor = &HFFC0FF
'                    Else
'                       .Col = wl_ColNo: .Row = (wY * 5) + 1
'                       .CellBackColor = &HC0FFFF
'                       .Col = wl_ColNo: .Row = (wY * 5) + 2
'                       .CellBackColor = &HC0FFFF
'                    End If
'                    If wDSP_Hac_Nai(wY, wX) = 1 Then
'                       .Col = wl_ColNo: .Row = (wY * 5) + 3
'                       .CellBackColor = &HFFC0FF
'                       .Col = wl_ColNo: .Row = (wY * 5) + 4
'                       .CellBackColor = &HFFC0FF
'                    Else
'                       .Col = wl_ColNo: .Row = (wY * 5) + 3
'                       .CellBackColor = &HC0FFFF
'                       .Col = wl_ColNo: .Row = (wY * 5) + 4
'                       .CellBackColor = &HC0FFFF
'                    End If
'
'                    .Row = wInt
                    
                    If wb_ZaiKB Then
                        Select Case wCMP_ZaiKb(wY, wX)
                            Case "1"
                            '   < 在庫の訂正 >
                                .TextMatrix((wY * 5) + 5, wl_ColNo) = "*" & Format(wl_Zaiko, "#####")
                            Case Else
                            '   < 通常 >
                                .TextMatrix((wY * 5) + 5, wl_ColNo) = Format(wl_Zaiko, "#####")
                        End Select
                        '在庫切れの場合は､色変え
                        If wl_Zaiko < 0 Then
                            wInt = .Row
                            .Col = wl_ColNo: .Row = (wY * 5) + 5
                            .CellBackColor = &HFFC0FF
                            .Row = wInt
                            
                        '===================================================================' 2019/08/05 ADD START
                        ' 在庫管理区分=3(発注点) かつ 在庫発注点＞在庫数　の場合は色変え
                        ElseIf wZK_KanriKbn = 3 And wZK_HacyuTem > wl_Zaiko Then    ' wl_Zaiko <> 0 And：ゼロを含まない場合はこの条件を使用する
                            wInt = .Row
                            .Col = wl_ColNo: .Row = (wY * 5) + 5
                            .CellBackColor = &HFFFF00
                            .Row = wInt
                        '===================================================================' 2019/08/05 ADD END
                        End If
                        
                    End If
                    
                
                End With
            End If
'            If wX = 21 Then
'               Debug.Print wX
'            End If
        Next wX
    Next wY
    
    vsGrid2.Redraw = True

End Sub
'+-----------------------------------------------+
'+      在庫履歴マスタ／材料在庫履歴マスタ - 在庫
'+-----------------------------------------------+
Private Function Zaiko_Set(Kubun As String, Optional DBNo As Integer = 0) As Boolean
    Dim ws_SetDate  As String   '>最終の実棚日

    Zaiko_Set = False

    '===============================================================' 2019/08/06 ADD START
    wZK_KanriKbn = 0        ' 在庫管理区分
    wZK_HacyuTem = 0        ' 在庫発注点
    '===============================================================' 2019/08/06 ADD END

'   >> 在庫管理区分
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  在庫管理区分"
    
    '===============================================================' 2019/08/06 ADD START
    Select Case Kubun
        Case "B": gSL_Select = gSL_Select & ", 在庫発注点"
        Case "Z": gSL_Select = gSL_Select & ", 在庫管理発注点"
    End Select
    '===============================================================' 2019/08/06 ADD END

    gSL_Select = gSL_Select & " from"
    Select Case Kubun
        Case "B": gSL_Select = gSL_Select & "  品番マスタ"
        Case "Z": gSL_Select = gSL_Select & "  材料マスタ"
    End Select
    Select Case Kubun
        Case "B": gSL_Select = gSL_Select & " where 品番 = '" & RTrim(cHinbn(wY)) & "'"
        Case "Z": gSL_Select = gSL_Select & " where 材料管理番号 = " & cZKanriNo(wY) & ""
    End Select
    
Call Str_Monitor(gSL_Select)

    If Not HNMRead(gSL_Select, 1, DBNo) Then Exit Function
    '===============================================================' 2019/08/06 ADD START
    Select Case Kubun
        Case "B":
            wZK_KanriKbn = HNM.在庫管理区分
            wZK_HacyuTem = HNM.在庫発注点
        Case "Z":
            wZK_KanriKbn = ZRM.在庫管理区分
            wZK_HacyuTem = ZRM.在庫管理発注点
    End Select
    '===============================================================' 2019/08/06 ADD END
    If HNM.在庫管理区分 = 0 Then Exit Function

'   >> 在庫ｾｯﾄ
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  年月度"
    gSL_Select = gSL_Select & ", 在庫数"
    gSL_Select = gSL_Select & ", 実棚日"
    gSL_Select = gSL_Select & " from"
    Select Case Kubun
        Case "B": gSL_Select = gSL_Select & "  在庫履歴マスタ"
        Case "Z": gSL_Select = gSL_Select & "  材料在庫履歴マスタ"
    End Select
    gSL_Select = gSL_Select & " where 年月度 between '" & CisFun.Mid2(wYmdS, 1, 6) & "' and '" & CisFun.Mid2(wYmdE, 1, 6) & "'"
    Select Case Kubun
'        Case "B": gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(cHinbn(wY)) & "'"
        Case "B"
                  gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(cHinbn(wY)) & "'"
                  If wSyoriDspKB = 3 Or wSyoriDspKB = 5 Then
                     gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(cTorcd(wY)) & "'"
                     gSL_Select = gSL_Select & "   and 工順 = " & cKoujyun(wY)
                  End If
        Case "Z": gSL_Select = gSL_Select & "   and 材料管理番号 = " & cZKanriNo(wY) & ""
    End Select
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  年月度 desc"
    
    If Kubun = "B" Then                         ' 2019/08/07 ADD
        gSL_Select = gSL_Select & ",  工順"          '''''2013/03/27
    End If                                      ' 2019/08/07 ADD
    
    gSL_Select = gSL_Select & ", 入力日付 desc"
    
    If RTrim(cHinbn(wY)) = "11058H1000" Then
        gInt = gInt
    End If
    
Call Str_Monitor(gSL_Select)

    If ZKRRead(gSL_Select, , DBNo) Then
        ws_SetDate = "99999999"
        Do Until ZKR_RDSTS = False
            With ZKR
                If .実棚日 = "" Then
                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(.年月度 & "01", "@@@@/@@/@@")) + 1
                    If wCMP_ZaiKb(wY, WDd) = "" Then
                        ws_SetDate = .年月度 & "01"
                        wCMP_ZaiSu(wY, WDd) = .在庫数
                        wCMP_ZaiKb(wY, WDd) = "0"
                    End If
                Else
                    If .実棚日 <= wDateE Then
                        If ws_SetDate > .実棚日 Then
                            ws_SetDate = .実棚日
                            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(.実棚日, "@@@@/@@/@@")) + 1
                            wCMP_ZaiSu(wY, WDd) = .在庫数
                            wCMP_ZaiKb(wY, WDd) = "1"
                        End If
                    End If
                End If
            End With
            Call ZKRReadNext(DBNo)
        Loop
    Else
    '   >> 在庫履歴マスタが無い時：最初の指示で作成した時の在庫数をｾｯﾄ
        If wFLG_YM = 1 Then wCMP_ZaiSu(wY, 0) = cZSu1(wY)     '>翌　月の時
        If wFLG_YM = 2 Then wCMP_ZaiSu(wY, 0) = cZSu2(wY)     '>翌々月の時
    End If
    Call ZKRClose(DBNo)

    Zaiko_Set = True
End Function
'****************************************************************
'*      納入指示構成ワーク作成      ･････取引先入力、品番未入力
'*           WSyoriDspKB = 0 or 1
'****************************************************************
Private Sub Naiji_Kosei_Crt0()
   Dim wLoop        As Integer
   Dim wMaxLoop     As Integer
   Dim wOLHinbn     As String
   Dim wIND         As Integer
   Dim wINDX        As String

   On Error GoTo Naiji_Kosei_Crt0_Err
' ( ワーク削除処理 )
    With CisDB
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        .SQL = "DELETE FROM 納入指示構成ワーク "
        .SQL = .SQL & " Where 処理端末 = SubString(Host_Name(),1,20)"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    End With

'   <><><><><> 登録取引先品番データ作成 <><><><><>
    If Select_HTM_HKT0(iH1_Torcd, 0) Then

    '======================== 2010/08/23 Start
'        Do Until Not HTM_RDSTS
'            Call ItemsClearNSK
'            With NSK
'                .処理端末 = "' + substring(host_name(), 1, 20) + '"
'                .順位 = ""
'                .レベル = 0
'                .部品材料区分 = "B"
'                If wSyoriDspKB = 0 Then                     '受注
'                   .大親品番 = RTrim(HTM.品番)
'                   .大親納入先 = RTrim(HTM.取引先)
'                Else                                        '発注
'                   .大親品番 = RTrim(HTM.品番)
'                   .大親納入先 = RTrim(HTM.取引先)
'                End If
'                .品番 = RTrim(HTM.品番)
'                .仕入先 = RTrim(HTM.取引先)
'                .受入 = RTrim(HTM.受入)
'            End With
'            Call NSKInsert
'
'            Call HTMReadNext
'        Loop
'
'        Call HTMClose

    End If

    Call Work_Set

    On Error GoTo 0

    Exit Sub
Naiji_Kosei_Crt0_Err:
    CisFun.ErrorBox
    End
End Sub
'+----------------------------------------------------+
'+      索引 - 品番取引先マスタ
'+----------------------------------------------------+
Private Function Select_HTM_HKT0(Torcd As String, Optional DBNo As Integer = 0) As Boolean
    Select_HTM_HKT0 = False

'   >> 品番取引先マスタ
    gSL_Select = ""
    
    '======================== 2010/08/23 Start
    
    gSL_Select = gSL_Select & " insert into 納入指示構成ワーク "
    gSL_Select = gSL_Select & " ( 処理端末,順位,レベル,部品材料区分,大親品番,大親納入先,品番,材料管理番号,"
    gSL_Select = gSL_Select & "   材質,板厚,幅,長さ,使用品番,仕入先,加工CD,展開区分,構成数,SEQNO,受入,作業順 )"
    
'    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  htm.品番"
'    gSL_Select = gSL_Select & ", htm.取引先"
'    gSL_Select = gSL_Select & ", htm.受入"
'    gSL_Select = gSL_Select & " from"
'    gSL_Select = gSL_Select & "  品番取引先マスタ htm"
'    gSL_Select = gSL_Select & " where htm.取引先 = '" & RTrim(Torcd) & "'"
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  host_name(),'',0,'B',htm.品番,htm.取引先,htm.品番,0,"
    gSL_Select = gSL_Select & "  '',0,0,0,'',htm.取引先,'',0,0,0,htm.受入,0    "
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  品番取引先マスタ htm"
    gSL_Select = gSL_Select & " left outer join 品番マスタ hn"
    gSL_Select = gSL_Select & " on  htm.品番 = hn.品番"
    
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create AlphaMe, "htm.取引先", RTrim(Torcd)
    End With
    If RTrim(iH2_Tokui) <> "" Then
        CisFun.Where_Create AlphaMe, "hn.得意先", iH2_Tokui
    End If
    If RTrim(iH2_Syakei) <> "" Then
        CisFun.Where_Create AlphaMe, "isnull(hn.車型,'')", iH2_Syakei
    End If
    If RTrim(iH2_Hinbn) <> "" Then
        CisFun.Where_Create AlphaMe, "htm.品番", iH2_Hinbn
    End If
        
    gSL_Select = gSL_Select & " " & CisFun.Where_Phrase
    '======================== 2010/08/23 End
    If H1Op_Torkb(0) Then
       gSL_Select = gSL_Select & "   and htm.品目 = 0"
    Else
       gSL_Select = gSL_Select & "   and htm.品目 = 1"
    End If
    
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
'    If Not HTMRead(gSL_Select, , DBNo) Then Exit Function


Call Str_Monitor(gSL_Select)

    With CisDB
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        .SQL = gSL_Select
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    End With
    
    
    Select_HTM_HKT0 = True
End Function
'*******************************************************
'*      納入指示構成ワーク作成････詳細情報表示(工順マスタ)
'*           WSyoriDspKB = 3 or 5
'*******************************************************
Private Sub Naiji_Kosei_Crt1()
   Dim wLoop        As Integer
   Dim wMaxLoop     As Integer
   Dim wOLHinbn     As String
   Dim wIND         As Integer
   Dim wINDX        As String

   On Error GoTo Naiji_Kosei_Crt1_Err
' ( ワーク削除処理 )
    With CisDB
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        .SQL = "DELETE FROM 納入指示構成ワーク "
        .SQL = .SQL & " Where 処理端末 = SubString(Host_Name(),1,20)"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    End With

'   <><><><><> レベル0データ作成 <><><><><>
    If Select_HTM_HKT1(iH1_Hinbn, 0) Then

        Do Until Not HKT_RDSTS
            Call ItemsClearNSK
            With NSK
                .処理端末 = "' + substring(host_name(), 1, 20) + '"
                .順位 = ""
                .レベル = 0
                .部品材料区分 = "B"
                .大親品番 = RTrim(iH1_Hinbn)
                .大親納入先 = RTrim(iH1_Torcd)
                .品番 = RTrim(HKT.品番)
                .仕入先 = RTrim(HKT.取引先)
                .加工CD = RTrim(HKT.機械)
        
                .受入 = RTrim(HKT.受入)
                .作業順 = HKT.工順
            End With
            Call NSKInsert
    
            Select Case NSK.展開区分
                '材料検索
                Case 0, 3:  If Not Insert_NSK_Zairyo(1) Then GoTo Naiji_Kosei_Crt1_Err
            End Select
            
            Call HKTReadNext
        Loop
        
        Call HKTClose

    End If

'   <><><><><> レベル1データ作成 <><><><><>
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  nsk.大親品番"
'    gSL_Select = gSL_Select & ", nsk.大親納入先"
'    gSL_Select = gSL_Select & ", kom.子品番"
'    gSL_Select = gSL_Select & ", kom.展開区分"
'    gSL_Select = gSL_Select & ", kom.構成数"
'    gSL_Select = gSL_Select & " from"
'    gSL_Select = gSL_Select & "  納入指示構成ワーク nsk"
'    gSL_Select = gSL_Select & ", 構成マスタ kom"
'    gSL_Select = gSL_Select & " where 処理端末 = substring(host_name(), 1, 20)"
'    gSL_Select = gSL_Select & "   and kom.親品番 = nsk.品番"
'    gSL_Select = gSL_Select & "   and nsk.部品材料区分 = 'B'"
'    gSL_Select = gSL_Select & "   and kom.展開区分 in (0, 1, 3)"
'    gSL_Select = gSL_Select & " order by"
'    gSL_Select = gSL_Select & "  nsk.順位"
'    gSL_Select = gSL_Select & ", kom.子品番"
'    If NSKRead(gSL_Select) Then
'        wIND = 100
'        Do Until NSK_RDSTS = False
'            If Select_HTM_HKT1(NSK.子品番, 1) Then
'
'                wIND = wIND + 1
'                With NSK
'                    .処理端末 = "' + substring(host_name(), 1, 20) + '"
'                    .順位 = Format(wIND, "000")
'                    .レベル = 1
'                    .部品材料区分 = "B"
'                    .大親品番 = RTrim(.大親品番)
'                    .大親納入先 = RTrim(.大親納入先)
'                    .品番 = RTrim(.子品番)
'                    .仕入先 = RTrim(HTM.取引先)
'                    .加工CD = RTrim(HKT.機械)
'                    .材料管理番号 = 0
'                    .板厚 = 0
'                    .幅 = 0
'                    .長さ = 0
'                    .使用品番 = ""
'
'                    .受入 = RTrim(HKT.受入)
'                    .作業順 = HKT.工順
'                End With
'                Call NSKInsert(1)
'
'                Select Case NSK.展開区分
'                    '材料検索
'                    Case 0, 3:  If Not Insert_NSK_Zairyo(1) Then GoTo Naiji_Kosei_Crt_Err
'                End Select
'            End If
'            Call NSKReadNext
'        Loop
'    End If
'    Call NSKClose

''   <><><><><> レベル2以降データ作成 <><><><><>
'    wMaxLoop = 99
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  nsk.大親品番"
'    gSL_Select = gSL_Select & ", nsk.大親納入先"
'    gSL_Select = gSL_Select & ", nsk.品番"
'    gSL_Select = gSL_Select & " from"
'    gSL_Select = gSL_Select & "  納入指示構成ワーク nsk"
'    gSL_Select = gSL_Select & " where nsk.処理端末 = substring(host_name(), 1, 20)"
'    gSL_Select = gSL_Select & "   and nsk.レベル = 1"
'    gSL_Select = gSL_Select & "   and nsk.部品材料区分 = 'B'"
'    gSL_Select = gSL_Select & "   and nsk.展開区分 in (0, 3)"
'    If NSKRead(gSL_Select) Then
'        wLoop = 1: wOLHinbn = ""
'        Do Until wLoop >= wMaxLoop
'            gSL_Select = ""
'            gSL_Select = gSL_Select & "select"
'            gSL_Select = gSL_Select & "  nsk.順位"
'            gSL_Select = gSL_Select & ", nsk.大親品番"
'            gSL_Select = gSL_Select & ", nsk.大親納入先"
'            gSL_Select = gSL_Select & ", kom.親品番"
'            gSL_Select = gSL_Select & ", kom.子品番"
'            gSL_Select = gSL_Select & ", kom.展開区分"
'            gSL_Select = gSL_Select & ", kom.構成数"
'            gSL_Select = gSL_Select & " from"
'            gSL_Select = gSL_Select & "  納入指示構成ワーク nsk"
'            gSL_Select = gSL_Select & ", 構成マスタ kom"
'            gSL_Select = gSL_Select & " where nsk.処理端末 = substring(host_name(), 1, 20)"
'            gSL_Select = gSL_Select & "   and nsk.レベル = " & wLoop
'            gSL_Select = gSL_Select & "   and nsk.部品材料区分 = 'B'"
'            gSL_Select = gSL_Select & "   and nsk.展開区分 in (0, 3)"
'            gSL_Select = gSL_Select & "   and kom.親品番 = nsk.品番"
'            gSL_Select = gSL_Select & "   and kom.展開区分 in (0, 1, 3)"
'            gSL_Select = gSL_Select & " order by"
'            gSL_Select = gSL_Select & "  nsk.順位"
'            gSL_Select = gSL_Select & ", kom.子品番"
'            If NSGRead(gSL_Select, , 1) Then
'                Do Until NSG_RDSTS = False
'                    If NSG.親品番 <> "" Then
'                        If wOLHinbn <> NSG.親品番 Then
'                            wIND = 100
'                            wOLHinbn = NSG.親品番
'                        End If
'
'                        wIND = wIND + 1
'                        wINDX = Format(wIND, "000")
'
'                        If Select_HTM_HKT1(NSG.子品番, 2) Then
'
'                            Call ItemsClearNSK
'                            With NSK
'                                .処理端末 = "' + substring(host_name(), 1, 20) + '"
'                                .順位 = NSG.順位 & wINDX
'                                .レベル = wLoop + 1
'                                .部品材料区分 = "B"
'                                .大親品番 = RTrim(NSG.大親品番)
'                                .大親納入先 = RTrim(NSG.大親納入先)
'                                .品番 = RTrim(NSG.子品番)
'                                .仕入先 = RTrim(HTM.取引先)
'                                .加工CD = RTrim(HKT.機械)
'                                .展開区分 = NSG.展開区分
'                                .構成数 = NSG.構成数
'
'                                .受入 = RTrim(HKT.受入)
'                                .作業順 = HKT.工順
'                            End With
'                            Call NSKInsert(2)
'
'                            Select Case NSK.展開区分
'                                '材料検索
'                                Case 0, 3:  If Not Insert_NSK_Zairyo(2) Then GoTo Naiji_Kosei_Crt_Err
'                            End Select
'
'                        End If
'                    End If
'                    Call NSGReadNext(1)
'                Loop
'            Else
'                wLoop = 98
'            End If
'            Call NSGClose(1)
'            wLoop = wLoop + 1
'            Call NSKReadNext
'        Loop
'
'    End If
'    Call NSKClose

    Call Work_Set

    On Error GoTo 0

    Exit Sub
Naiji_Kosei_Crt1_Err:
    CisFun.ErrorBox
    End
End Sub
'+----------------------------------------------------+
'+      索引 - 品番取引先マスタ～品番工順マスタ
'+----------------------------------------------------+
Private Function Select_HTM_HKT1(Hinbn As String, Optional DBNo As Integer = 0) As Boolean
    Select_HTM_HKT1 = False

'   >> 品番取引先マスタ
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  htm.品番"
'    gSL_Select = gSL_Select & ", htm.取引先"
'    gSL_Select = gSL_Select & " from"
'    gSL_Select = gSL_Select & "  品番取引先マスタ htm"
'    gSL_Select = gSL_Select & " where htm.品番 = '" & RTrim(Hinbn) & "'"
'    gSL_Select = gSL_Select & "   and htm.品目 = 1"
'    gSL_Select = gSL_Select & " order by"
'    gSL_Select = gSL_Select & "  htm.取引先"
'    gSL_Select = gSL_Select & ", htm.受入"
'    If Not HTMRead(gSL_Select, 1, DBNo) Then Exit Function


'   >> 品番工順マスタ
    Call ItemsClearHKT
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  HKT.機械"
    gSL_Select = gSL_Select & "  HKT.品番,HKT.取引先, "
    gSL_Select = gSL_Select & "  HKT.機械, HKT.工順, HKT.受入"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  品番工順マスタ HKT"
    gSL_Select = gSL_Select & " where HKT.品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   and HKT.取引先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & " order by"
'    gSL_Select = gSL_Select & "  HKT.機械順 desc, HKT.受入"
    gSL_Select = gSL_Select & "  HKT.工順 desc, HKT.受入"
    Call HKTRead(gSL_Select, , DBNo)

    Select_HTM_HKT1 = True
End Function
'*******************************************************
'*      納入指示構成ワーク作成
'*******************************************************
Private Sub Naiji_Kosei_Crt2()
   Dim wLoop        As Integer
   Dim wMaxLoop     As Integer
   Dim wOLHinbn     As String
   Dim wIND         As Integer
   Dim wINDX        As String

   On Error GoTo Naiji_Kosei_Crt2_Err
' ( ワーク削除処理 )
    With CisDB
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        .SQL = "DELETE FROM 納入指示構成ワーク "
        .SQL = .SQL & " Where 処理端末 = SubString(Host_Name(),1,20)"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    End With

'    If wSyoriDspKB = 4 Then
'       Call GyakuTenkai(RTrim(iH1_Hinbn), 0)
'    End If

'   <><><><><> レベル0データ作成 <><><><><>
'    If Select_HTM_HKT2(iH1_Hinbn, 0) Then
    If Select_HTM_HKT2_1(wOyaHinbn, 0) Then

        Call ItemsClearNSK
        With NSK
            .処理端末 = "' + substring(host_name(), 1, 20) + '"
            .順位 = ""
            .レベル = 0
            .部品材料区分 = "B"
            '.大親品番 = RTrim(iH1_Hinbn)
            '.大親納入先 = RTrim(iH1_Torcd)
            .大親品番 = RTrim(wOyaHinbn)
'            .大親納入先 = RTrim(wOyaTorcd)
            If wSyoriDspKB <> 7 Then
                .大親納入先 = RTrim(wOyaTorcd)
            Else
                .大親納入先 = RTrim(HTM.取引先)
            End If
            .品番 = RTrim(HTM.品番)
            .仕入先 = RTrim(HTM.取引先)
            .加工CD = ""
    
            .受入 = RTrim(HTM.受入)
            .作業順 = 0
        End With
        Call NSKInsert

        Select Case NSK.展開区分
            '材料検索
            Case 0, 3:  If Not Insert_NSK_Zairyo(1) Then GoTo Naiji_Kosei_Crt2_Err
        End Select

    End If

'   <><><><><> レベル1データ作成 <><><><><>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  nsk.大親品番"
    gSL_Select = gSL_Select & ", nsk.大親納入先"
    gSL_Select = gSL_Select & ", kom.子品番"
    gSL_Select = gSL_Select & ", kom.展開区分"
    gSL_Select = gSL_Select & ", kom.構成数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  納入指示構成ワーク nsk"
    gSL_Select = gSL_Select & ", 構成マスタ kom"
    gSL_Select = gSL_Select & " where 処理端末 = substring(host_name(), 1, 20)"
    gSL_Select = gSL_Select & "   and kom.親品番 = nsk.品番"
    gSL_Select = gSL_Select & "   and nsk.部品材料区分 = 'B'"
    gSL_Select = gSL_Select & "   and kom.展開区分 in (0, 1, 3)"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  nsk.順位"
    gSL_Select = gSL_Select & ", kom.子品番"
    If NSKRead(gSL_Select) Then
        wIND = 100
        Do Until NSK_RDSTS = False
            If Select_HTM_HKT2_2(NSK.子品番, 1) Then

                wIND = wIND + 1
                With NSK
                    .処理端末 = "' + substring(host_name(), 1, 20) + '"
                    .順位 = Format(wIND, "000")
                    .レベル = 1
                    .部品材料区分 = "B"
                    .大親品番 = RTrim(.大親品番)
                    .大親納入先 = RTrim(.大親納入先)
                    .品番 = RTrim(.子品番)
                    .仕入先 = RTrim(HTM.取引先)
                    .加工CD = ""
                    .材料管理番号 = 0
                    .板厚 = 0
                    .幅 = 0
                    .長さ = 0
                    .使用品番 = ""
    
                    .受入 = RTrim(HTM.受入)
                    .作業順 = 0
                End With
                Call NSKInsert(1)

                Select Case NSK.展開区分
                    '材料検索
                    Case 0, 3:  If Not Insert_NSK_Zairyo(1) Then GoTo Naiji_Kosei_Crt2_Err
                End Select
            End If
            Call NSKReadNext
        Loop
    End If
    Call NSKClose

'   <><><><><> レベル2以降データ作成 <><><><><>
    wMaxLoop = 99
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  nsk.大親品番"
    gSL_Select = gSL_Select & ", nsk.大親納入先"
    gSL_Select = gSL_Select & ", nsk.品番"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  納入指示構成ワーク nsk"
    gSL_Select = gSL_Select & " where nsk.処理端末 = substring(host_name(), 1, 20)"
    gSL_Select = gSL_Select & "   and nsk.レベル = 1"
    gSL_Select = gSL_Select & "   and nsk.部品材料区分 = 'B'"
    gSL_Select = gSL_Select & "   and nsk.展開区分 in (0, 3)"
    If NSKRead(gSL_Select) Then
        wLoop = 1: wOLHinbn = ""
        Do Until wLoop >= wMaxLoop
            gSL_Select = ""
            gSL_Select = gSL_Select & "select"
            gSL_Select = gSL_Select & "  nsk.順位"
            gSL_Select = gSL_Select & ", nsk.大親品番"
            gSL_Select = gSL_Select & ", nsk.大親納入先"
            gSL_Select = gSL_Select & ", kom.親品番"
            gSL_Select = gSL_Select & ", kom.子品番"
            gSL_Select = gSL_Select & ", kom.展開区分"
            gSL_Select = gSL_Select & ", kom.構成数"
            gSL_Select = gSL_Select & " from"
            gSL_Select = gSL_Select & "  納入指示構成ワーク nsk"
            gSL_Select = gSL_Select & ", 構成マスタ kom"
            gSL_Select = gSL_Select & " where nsk.処理端末 = substring(host_name(), 1, 20)"
            gSL_Select = gSL_Select & "   and nsk.レベル = " & wLoop
            gSL_Select = gSL_Select & "   and nsk.部品材料区分 = 'B'"
            gSL_Select = gSL_Select & "   and nsk.展開区分 in (0, 3)"
            gSL_Select = gSL_Select & "   and kom.親品番 = nsk.品番"
            gSL_Select = gSL_Select & "   and kom.展開区分 in (0, 1, 3)"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  nsk.順位"
            gSL_Select = gSL_Select & ", kom.子品番"
            If NSGRead(gSL_Select, , 1) Then
                Do Until NSG_RDSTS = False
                    If NSG.親品番 <> "" Then
                        If wOLHinbn <> NSG.親品番 Then
                            wIND = 100
                            wOLHinbn = NSG.親品番
                        End If

                        wIND = wIND + 1
                        wINDX = Format(wIND, "000")

                        If Select_HTM_HKT2_2(NSG.子品番, 2) Then

                            Call ItemsClearNSK
                            With NSK
                                .処理端末 = "' + substring(host_name(), 1, 20) + '"
                                .順位 = NSG.順位 & wINDX
                                .レベル = wLoop + 1
                                .部品材料区分 = "B"
                                .大親品番 = RTrim(NSG.大親品番)
                                .大親納入先 = RTrim(NSG.大親納入先)
                                .品番 = RTrim(NSG.子品番)
                                .仕入先 = RTrim(HTM.取引先)
                                .加工CD = RTrim(HKT.機械)
                                .展開区分 = NSG.展開区分
                                .構成数 = NSG.構成数
    
                                .受入 = RTrim(HKT.受入)
                                .作業順 = HKT.工順
                            End With
                            Call NSKInsert(2)

                            Select Case NSK.展開区分
                                '材料検索
                                Case 0, 3:  If Not Insert_NSK_Zairyo(2) Then GoTo Naiji_Kosei_Crt2_Err
                            End Select
                   
                        End If
                    End If
                    Call NSGReadNext(1)
                Loop
            Else
                wLoop = 98
            End If
            Call NSGClose(1)
            wLoop = wLoop + 1
            Call NSKReadNext
        Loop

    End If
    Call NSKClose

    Call Work_Set

    On Error GoTo 0

    Exit Sub
Naiji_Kosei_Crt2_Err:
    CisFun.ErrorBox
    End
End Sub
'+----------------------------------------------------+
'+      索引 - 品番取引先マスタ----(大親品番)
'+----------------------------------------------------+
Private Function Select_HTM_HKT2_1(Hinbn As String, Optional DBNo As Integer = 0) As Boolean
    Select_HTM_HKT2_1 = False

'   >> 品番取引先マスタ(大親)
    If wSyoriDspKB <> 7 Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  htm.品番"
        gSL_Select = gSL_Select & ", htm.取引先"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  品番取引先マスタ htm"
        gSL_Select = gSL_Select & " where htm.品番 = '" & RTrim(Hinbn) & "'"
        gSL_Select = gSL_Select & "   and htm.品目 = 0"
        gSL_Select = gSL_Select & " order by"
        gSL_Select = gSL_Select & "  htm.取引先"
        gSL_Select = gSL_Select & ", htm.受入"
        If Not HTMRead(gSL_Select, 1, DBNo) Then Exit Function
        
        wOyaTorcd = HTM.取引先
    End If
    
'   >> 品番取引先マスタ
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  htm.品番"
    gSL_Select = gSL_Select & ", htm.取引先"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  品番取引先マスタ htm"
    gSL_Select = gSL_Select & " where htm.品番 = '" & RTrim(Hinbn) & "'"
    gSL_Select = gSL_Select & "   and htm.品目 = 1"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    If Not HTMRead(gSL_Select, 1, DBNo) Then Exit Function
    

'   >> 品番工順マスタ
'    Call ItemsClearHKT
'
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  HKT.機械, HKT.工順, HKT.受入"
'    gSL_Select = gSL_Select & " from"
'    gSL_Select = gSL_Select & "  品番工順マスタ HKT"
'    gSL_Select = gSL_Select & " where HKT.品番 = '" & RTrim(HTM.品番) & "'"
'    gSL_Select = gSL_Select & "   and HKT.取引先 = '" & RTrim(HTM.取引先) & "'"
'    gSL_Select = gSL_Select & " order by"
'    gSL_Select = gSL_Select & "  HKT.工順 desc, HKT.受入"
'    Call HKTRead(gSL_Select, 1, DBNo)

    Select_HTM_HKT2_1 = True
End Function
'+----------------------------------------------------+
'+      索引 - 品番取引先マスタ～品番工順マスタ
'+----------------------------------------------------+
Private Function Select_HTM_HKT2_2(Hinbn As String, Optional DBNo As Integer = 0) As Boolean
    Select_HTM_HKT2_2 = False

'   >> 品番取引先マスタ
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  htm.品番"
    gSL_Select = gSL_Select & ", htm.取引先"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  品番取引先マスタ htm"
    gSL_Select = gSL_Select & " where htm.品番 = '" & RTrim(Hinbn) & "'"
    gSL_Select = gSL_Select & "   and htm.品目 = 1"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    If Not HTMRead(gSL_Select, 1, DBNo) Then Exit Function
    

'   >> 品番工順マスタ
'    Call ItemsClearHKT
'
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  HKT.機械, HKT.工順, HKT.受入"
'    gSL_Select = gSL_Select & " from"
'    gSL_Select = gSL_Select & "  品番工順マスタ HKT"
'    gSL_Select = gSL_Select & " where HKT.品番 = '" & RTrim(HTM.品番) & "'"
'    gSL_Select = gSL_Select & "   and HKT.取引先 = '" & RTrim(HTM.取引先) & "'"
'    gSL_Select = gSL_Select & " order by"
'    gSL_Select = gSL_Select & "  HKT.工順 desc, HKT.受入"
'    Call HKTRead(gSL_Select, 1, DBNo)

    Select_HTM_HKT2_2 = True
End Function
'**************************************************************
'*      大親品番の取得 - 逆展開     [構成/親品番]
'*          親品番単一 = 0 / 親品番複数 = 1
'**************************************************************
Private Function GyakuTenkai(KHinbn As String, Optional DBNo As Integer = 1) As Byte

    With CisDB
        .SQL = "構成問合せ情報抽出"
        .StoadoCount = 4
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
        .ParaNo = 4: .ParaIO = Input用
    
        .ParaNo = 1: .ParaValue = RTrim(KHinbn)
        .ParaNo = 2: .ParaValue = 0         '0:部品    1:材料
        .ParaNo = 3: .ParaValue = 1         '0:正展開  1:逆展開
        .ParaNo = 4: .ParaValue = 1         '0:直親    1:大親
    
        If Not .DBStored Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     構成マスタ情報の抽出に失敗しました。     "
                .MB_Title = "抽出処理"
                .MB_Button = Error
                If .MBOX Then Exit Function
            End With
        End If
    End With

    gSL_Select = "SELECT COUNT(*) 件数 FROM 構成問合せワーク KW "
    gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
    
    If Not KOLRead(gSL_Select, , 1) Then
        Call KOLClose(DBNo)
        Exit Function
    Else
        If KOL.件数 > 1 Then
           GyakuTenkai = 1
        Else
           GyakuTenkai = 0
           wOyaHinbn = KHinbn
        End If
        Call KOLClose(DBNo)
    End If

    If GyakuTenkai = 0 Then
       gSL_Select = "SELECT 親品番,子品番 FROM 構成問合せワーク "
       gSL_Select = gSL_Select & " WHERE 処理端末 = HOST_NAME()"
       If KOLRead(gSL_Select, 1, 1) Then
          wOyaHinbn = KOL.親品番
       Else
           wOyaHinbn = KHinbn
       End If
    End If
    
End Function
'+-------------------------------------+
'+      納入指示構成ワーク作成（材料）
'+-------------------------------------+
Private Function Insert_NSK_Zairyo(Optional DBNo As Integer = 0) As Boolean
    Dim wb_Count    As Byte
    Dim ws_Jyuni    As String   '>順位
    Dim wb_Level    As Byte     '>レベル

    Insert_NSK_Zairyo = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  zhm.材料管理番号"
    gSL_Select = gSL_Select & ", zhm.使用量"
    gSL_Select = gSL_Select & ", isnull(zrm.材質, '') 材質"
    gSL_Select = gSL_Select & ", isnull(zrm.板厚, 0) 板厚"
    gSL_Select = gSL_Select & ", isnull(zrm.幅, 0) 幅"
    gSL_Select = gSL_Select & ", isnull(zrm.長さ, 0) 長さ"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料品番マスタ zhm"
    gSL_Select = gSL_Select & "  left outer join 材料マスタ zrm"
    gSL_Select = gSL_Select & "    on zrm.材料管理番号 = zhm.材料管理番号"
    gSL_Select = gSL_Select & " where zhm.使用品番 = '" & RTrim(NSK.品番) & "'"
    gSL_Select = gSL_Select & "   and zhm.展開区分 = 0"
    If ZHMRead(gSL_Select, , 3) Then
        wb_Count = 0
        ws_Jyuni = NSK.順位
        wb_Level = NSK.レベル
        Do Until Not ZHM_RDSTS
            wb_Count = wb_Count + 1
            With NSK
            '   .処理端末 = "' + substring(host_name(), 1, 20) + '"
                .順位 = ws_Jyuni & Format(wb_Count, "000")
                .レベル = wb_Level + 1
                .部品材料区分 = "Z"
            '   .大親品番 = .大親品番
            '   .大親納入先 = .大親納入先
            '   .品番 = .品番
            '   .仕入先 = .仕入先
                .材料管理番号 = ZHM.材料管理番号
                .材質 = ZHM.材質
                .板厚 = ZHM.板厚
                .幅 = ZHM.幅
                .長さ = ZHM.長さ
                .使用品番 = .品番
                .加工CD = ""
                .展開区分 = 0
                .構成数 = ZHM.使用量
            End With
            Call NSKInsert(DBNo)
            Call ZHMReadNext(3)
        Loop
    End If
    Call ZHMClose(3)

    Insert_NSK_Zairyo = True
End Function
'*************************************************************
'* ワーク項目にセット
'*************************************************************
Private Sub Work_Set()
    Erase wDSP_JyuSu
    Erase wDSP_HacSu
    Erase wDSP_SyuSu
    Erase wDSP_NyuSu
    Erase wCMP_ZaiSu
    Erase wCMP_ZaiKb
    Erase wCMP_NyuSu
    Erase wCMP_SyuSu
    Erase cHinbn
    Erase cZKanriNo
    Erase cTorcd
    Erase cUkeir
    Erase cKoujyun
    
    Erase wDSP_Jyu_Nai
    Erase wDSP_Hac_Nai

    wSyuSu = 0: wZenSyuSu = 0
    
    With CisFun
    '   >> 開始日：前月／終了１０日間
        .Date_RsFormat = fYMD
        .Date_Add wDSP_YM & "01", D, -10
        wDateS = .Date_Result
    '   >> 終了日：翌月／開始１０日間
        .Date_RsFormat = fYM
        .Date_Add wDSP_YM & "01", M, 1
        wDateE = .Date_Result & "10"
    End With
'   >> 抽出用の日付ｾｯﾄ
    wYmdS = CisFun.Mid2(wDateS, 1, 6) & "01"
    wYmdE = wDateE

    Call wCLM_Chk           'カレンダマスタチェック
    
    wY = 0: wX = 0: wKensu = 0
    gSL_Select = "SELECT * FROM 納入指示構成ワーク "
    gSL_Select = gSL_Select & " WHERE 処理端末 = SubString(Host_Name(),1,20)"
    gSL_Select = gSL_Select & "   AND 展開区分 in (0, 1)"
    If wSyoriDspKB = 3 Or wSyoriDspKB = 5 Then
       gSL_Select = gSL_Select & " ORDER BY 順位,品番,部品材料区分,作業順"
    Else
       gSL_Select = gSL_Select & " ORDER BY 順位,品番,部品材料区分"
    End If
    
    
Call Str_Monitor(gSL_Select)
    
    
    If NSKRead(gSL_Select, , 0) Then
        Do Until NSK_RDSTS = False
            If NSK.レベル = 0 Then      '大親処理
                If wSyoriDspKB = 0 Or wSyoriDspKB = 4 Then
                    If H2Op_DataKB(0) Then wDSP_Flg = 0             '2010/08/23
                    '取引先単位：受注
                    wHinbn = RTrim(NSK.大親品番)
                    wTorcd = RTrim(NSK.大親納入先)
                    Call Kansei_JYT_Set             '<受注テーブル>     出荷予定
                    
'====================== 2010/08/23 Start
'                    Call Buhin_JKT_Set(2)           '<実績管理テーブル> 出荷実績
'                    Call Kansei_SKT_Set(wTorcd)     '<出荷テーブル>     出荷実績
'                    wHinbn = RTrim(NSK.品番)
'                    wTorcd = RTrim(NSK.仕入先)
'                    Call Kansei_HCT_Set             '<発注テーブル>     入荷予定
'                    Call Buhin_JKT_Set(1)           '<実績管理テーブル> 入荷実績
'                    Call Kyotu_NYT_Set              '<入荷テーブル>     入荷実績
'                    cHinbn(wY) = RTrim(NSK.品番)
'                    wY = wY + 1                     '''''''
'                    wKensu = wKensu + 1             '''''''

                    If wSyoriDspKB = 0 And H2Op_DataKB(0).Value = True Then
                        If wDSP_Flg = 1 Then
                            Call Buhin_JKT_Set(2)           '<実績管理テーブル> 出荷実績
                            Call Kansei_SKT_Set(wTorcd)     '<出荷テーブル>     出荷実績
                            wHinbn = RTrim(NSK.品番)
                            wTorcd = RTrim(NSK.仕入先)
                            Call Kansei_HCT_Set             '<発注テーブル>     入荷予定
                            Call Buhin_JKT_Set(1)           '<実績管理テーブル> 入荷実績
                            Call Kyotu_NYT_Set              '<入荷テーブル>     入荷実績
                            cHinbn(wY) = RTrim(NSK.品番)
                            wY = wY + 1                     '''''''
                            wKensu = wKensu + 1             '''''''
                        Else
                            With CisDB
                                .SQL = "DELETE FROM 納入指示構成ワーク "
                                .SQL = .SQL & " Where 処理端末 = SubString(Host_Name(),1,20)"
                                .SQL = .SQL & " And   大親品番 = '" & RTrim(NSK.大親品番) & "'"
                                .SQL = .SQL & " And   品番 = '" & RTrim(NSK.品番) & "'"
                                .DBExec
                            End With
                        End If
                    Else
                        Call Buhin_JKT_Set(2)           '<実績管理テーブル> 出荷実績
                        Call Kansei_SKT_Set(wTorcd)     '<出荷テーブル>     出荷実績
                        wHinbn = RTrim(NSK.品番)
                        wTorcd = RTrim(NSK.仕入先)
                        Call Kansei_HCT_Set             '<発注テーブル>     入荷予定
                        Call Buhin_JKT_Set(1)           '<実績管理テーブル> 入荷実績
                        Call Kyotu_NYT_Set              '<入荷テーブル>     入荷実績
                        cHinbn(wY) = RTrim(NSK.品番)
                        cTorcd(wY) = RTrim(iH1_Torcd)
                        wY = wY + 1                     '''''''
                        wKensu = wKensu + 1             '''''''
                    End If
                End If
'====================== 2010/08/23 End
                If wSyoriDspKB = 1 Then
                    If H2Op_DataKB(0) Then wDSP_Flg = 0             '2010/08/23
                    '取引先単位：発注
                    wHinbn = RTrim(NSK.大親品番)
                    Call GyakuTenkai_Naiji(wHinbn)  '<内示テーブル(大親品番)>     出荷予定
                    
'=========================== 2010/08/23 Start
'                    Call Buhin_JKT_Set              '<実績管理テーブル> 出庫実績
'                    Call Kansei_HCT_Set             '<発注テーブル>     入荷予定
'                    Call Kansei_SKT_Set             '<出荷テーブル>     出荷実績
'                    Call Kyotu_NYT_Set              '<入荷テーブル>     入荷実績 ........ 手配先単位の時は: (wTorcd) を指定
'                    cHinbn(wY) = RTrim(NSK.品番)
'                    wY = wY + 1                     '''''''
'                    wKensu = wKensu + 1             '''''''
                    If H2Op_DataKB(0) Then
                        If wDSP_Flg = 1 Then
                            Call Buhin_JKT_Set              '<実績管理テーブル> 出庫実績
                            Call Kansei_HCT_Set             '<発注テーブル>     入荷予定
                            Call Kansei_SKT_Set             '<出荷テーブル>     出荷実績
                            Call Kyotu_NYT_Set              '<入荷テーブル>     入荷実績 ........ 手配先単位の時は: (wTorcd) を指定
                            cHinbn(wY) = RTrim(NSK.品番)
                            cTorcd(wY) = RTrim(iH1_Torcd)                   '2010/11/11
                            wY = wY + 1                     '''''''
                            wKensu = wKensu + 1             '''''''
                        Else
                            With CisDB
                                .SQL = "DELETE FROM 納入指示構成ワーク "
                                .SQL = .SQL & " Where 処理端末 = SubString(Host_Name(),1,20)"
                                .SQL = .SQL & " And   大親品番 = '" & RTrim(NSK.大親品番) & "'"
                                .SQL = .SQL & " And   品番 = '" & RTrim(NSK.品番) & "'"
                                .DBExec
                            End With
                        End If
                    Else
                        Call Buhin_JKT_Set              '<実績管理テーブル> 出庫実績
                        Call Kansei_HCT_Set             '<発注テーブル>     入荷予定
                        Call Kansei_SKT_Set             '<出荷テーブル>     出荷実績
                        Call Kyotu_NYT_Set              '<入荷テーブル>     入荷実績 ........ 手配先単位の時は: (wTorcd) を指定
                        cHinbn(wY) = RTrim(NSK.品番)
                        cTorcd(wY) = RTrim(iH1_Torcd)                   '2010/11/11
                        wY = wY + 1                     '''''''
                        wKensu = wKensu + 1             '''''''
                    End If
'=========================== 2010/08/23 End
                
                End If
                If wSyoriDspKB = 3 Or wSyoriDspKB = 5 Then
                    '取引先/品番単位：詳細
                    wHinbn = RTrim(NSK.大親品番)
                    Call GyakuTenkai_Naiji(wHinbn)  '<内示テーブル(大親品番)>     出荷予定
                    Call Buhin_JKT_Set              '<実績管理テーブル> 出庫実績
                    Call Kansei_HCT_Set             '<発注テーブル>     入荷予定
                    Call Kansei_SKT_Set             '<出荷テーブル>     出荷実績
                    Call Kyotu_NYT_Set              '<入荷テーブル>     入荷実績 ........ 手配先単位の時は: (wTorcd) を指定
                    cHinbn(wY) = RTrim(NSK.品番)
                    cTorcd(wY) = RTrim(NSK.仕入先)
                    cKoujyun(wY) = RTrim(NSK.作業順)
                    wY = wY + 1                     '''''''
                    wKensu = wKensu + 1             '''''''
                End If
                If wSyoriDspKB = 2 Or wSyoriDspKB = 6 Then
                    '取引先/品番単位：受注/構成  or  受注/指定品番以下
                    '取引先単位：受注
                    wHinbn = RTrim(NSK.大親品番)
                    wTorcd = RTrim(NSK.大親納入先)
                    Call Kansei_JYT_Set             '<受注テーブル>     出荷予定
                    Call Buhin_JKT_Set(2)           '<実績管理テーブル> 出荷実績
                    Call Kansei_SKT_Set(wTorcd)     '<出荷テーブル>     出荷実績
                    wHinbn = RTrim(NSK.品番)
                    wTorcd = RTrim(NSK.仕入先)
                    Call Kansei_HCT_Set             '<発注テーブル>     入荷予定
                    Call Buhin_JKT_Set(1)           '<実績管理テーブル> 入荷実績
                    Call Kyotu_NYT_Set              '<入荷テーブル>     入荷実績
                    cHinbn(wY) = RTrim(NSK.品番)
                    cTorcd(wY) = RTrim(NSK.仕入先)                      '2010/11/11
                    wY = wY + 1                     '''''''
                    wKensu = wKensu + 1             '''''''
                End If
                If wSyoriDspKB = 7 Then
                    '取引先/品番単位：発注/構成  or  発注/指定品番以下
                    '取引先単位：発注
                    wHinbn = RTrim(NSK.大親品番)
                    wTorcd = RTrim(NSK.大親納入先)
                    Call Kansei_JYT_Set             '<受注テーブル>     出荷予定
                    Call Buhin_JKT_Set(2)           '<実績管理テーブル> 出荷実績
                    Call Kansei_SKT_Set(wTorcd)     '<出荷テーブル>     出荷実績
                    wHinbn = RTrim(NSK.品番)
                    wTorcd = RTrim(NSK.仕入先)
                    Call Kansei_HCT_Set             '<発注テーブル>     入荷予定
                    Call Buhin_JKT_Set(1)           '<実績管理テーブル> 入荷実績
                    Call Kyotu_NYT_Set              '<入荷テーブル>     入荷実績
                    cHinbn(wY) = RTrim(NSK.品番)
                    cTorcd(wY) = RTrim(NSK.仕入先)                      '2010/11/11
                    wY = wY + 1                     '''''''
                    wKensu = wKensu + 1             '''''''
                End If
            Else
                If NSK.部品材料区分 = "B" Then
                    wHinbn = RTrim(NSK.品番)
                    wTorcd = RTrim(NSK.仕入先)
'                    wY = wY + 1
                  '>Call Buhin_HNJ_Set          '<発注内示テーブル> 入荷予定 ........ 手配先単位の時は: (wTorcd) を指定
                    Call GyakuTenkai_Naiji(wHinbn)  '<内示テーブル(大親品番)>     出荷予定
                    Call Kansei_HCT_Set         '<発注テーブル>     入荷予定
                    Call Buhin_JKT_Set          '<実績管理テーブル> 出荷実績＆入荷実績
                    Call Kansei_SKT_Set         '<出荷テーブル>     出荷実績
                    Call Kyotu_NYT_Set          '<入荷テーブル>     入荷実績 ........ 手配先単位の時は: (wTorcd) を指定
                    cHinbn(wY) = RTrim(NSK.品番)
                    wY = wY + 1
                Else
                    wZaisitu = RTrim(NSK.材質)
                    wItaatu = NSK.板厚
                    wHaba = NSK.幅
                    wNagasa = NSK.長さ
                    wZKanriNo = NSK.材料管理番号
'                    wY = wY + 1
                  '>Call Zairyo_ZNJ_Set         '<材料発注内示テーブル> 入荷予定
                    Call Zairyo_ZCT_Set         '<材料発注テーブル>     入荷予定
                    Call Zairyo_ZJK_Set         '<材料実績管理テーブル> 出荷実績＆入荷実績
                    Call Zairyo_ZYT_Set         '<材料入荷テーブル>     入荷実績
                    cZKanriNo(wY) = RTrim(NSK.材料管理番号)
                    wY = wY + 1
                End If
                wKensu = wKensu + 1
            End If
            Call NSKReadNext(0)
        Loop
    End If
End Sub
'**************************************************************
'*      大親品番の取得/内示情報の取得 - 逆展開
'**************************************************************
Private Sub GyakuTenkai_Naiji(KHinbn As String, Optional DBNo As Integer = 1)

'======================================== 2010/09/01 Start
'    With CisDB
'        .SQL = "構成問合せ情報抽出"
'        .StoadoCount = 4
'        .ParaNo = 0: .ParaIO = Return用
'        .ParaNo = 1: .ParaIO = Input用
'        .ParaNo = 2: .ParaIO = Input用
'        .ParaNo = 3: .ParaIO = Input用
'        .ParaNo = 4: .ParaIO = Input用
'
'        .ParaNo = 1: .ParaValue = RTrim(KHinbn)
'        .ParaNo = 2: .ParaValue = 0         '0:部品    1:材料
'        .ParaNo = 3: .ParaValue = 1         '0:正展開  1:逆展開
'        .ParaNo = 4: .ParaValue = 1         '0:直親    1:大親              '2010/09/01
'
'        If Not .DBStored Then
'            With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "     構成マスタ情報の抽出に失敗しました。     "
'                .MB_Title = "抽出処理"
'                .MB_Button = Error
'                If .MBOX Then Exit Sub
'            End With
'        End If
'    End With
    If Not Kosei_Toiawase(KHinbn, 1, 1) Then Exit Sub           '大親品番の取得
'======================================== 2010/09/01 End

    gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示品番 FROM 構成問合せワーク KW "
    gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
    gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示品番"
    gSL_Select = gSL_Select & " ORDER BY KW.親品番"
    
    If Not KOLRead(gSL_Select, , 1) Then
        Call KOLClose(1)
                
'        '構成に無い場合は、入力品番にてチェック
'        Call Oya_NJT_Set(RTrim(iH1_Hinbn), 2)
        
        '構成に無い場合
        If RTrim(iH1_Hinbn) <> "" Then
           '個別の場合は､入力品番にてチェック
           Call Oya_NJT_Set(RTrim(iH1_Hinbn), 2)
        Else
           '取引先単位の場合は､親品番にてチェック
           Call Oya_NJT_Set(RTrim(KHinbn), 2)
        End If
        
        wNYMD = ""          '2010/09/01
        Exit Sub
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "     構成マスタ情報が存在しません。     "
'            .MB_Title = "抽出処理"
'            .MB_Button = Error
'            If .MBOX Then Exit Function
'        End With
    Else
        Do Until Not KOL_RDSTS
            Call KoseiSu_Get(KOL.親品番, KHinbn)            '2010/09/01
            Call Oya_NJT_Set(KOL.親品番, 2)
            Call KOLReadNext(1)
        Loop
        Call KoseiSu_Get(KHinbn, KHinbn)                    '2010/09/01
        Call Oya_NJT_Set(RTrim(KHinbn), 2)
        Call KOLClose(1)
        wNYMD = ""          '2010/09/01
    End If

End Sub
'**************************************************************
'*      逆展開-構成問合せ           2010/09/01
'**************************************************************
Private Function Kosei_Toiawase(Kohinbn As String, TenkaiKB As Byte, OyaKB As Byte) As Boolean

    Kosei_Toiawase = False
    
    With CisDB
        .SQL = "構成問合せ情報抽出"
        .StoadoCount = 4
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
        .ParaNo = 4: .ParaIO = Input用
    
        .ParaNo = 1: .ParaValue = RTrim(Kohinbn)
        .ParaNo = 2: .ParaValue = 0         '0:部品    1:材料
        .ParaNo = 3: .ParaValue = TenkaiKB  '0:正展開  1:逆展開
        .ParaNo = 4: .ParaValue = OyaKB     '0:直親    1:大親
        
        If Not .DBStored Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     構成マスタ情報の抽出に失敗しました。     "
                .MB_Title = "抽出処理"
                .MB_Button = Error
                If .MBOX Then Exit Function
            End With
        End If
    End With

    Kosei_Toiawase = True
    
End Function
'**************************************************************
'*      構成数取得                             2010/09/01
'**************************************************************
Private Function KoseiSu_Get(KHinbn As String, Kohinbn As String) As Boolean
    KoseiSu_Get = False
    wKoseiSu = 1
    
    With CisDB
        .SQL = "構成マスタメンテ情報抽出"
        .StoadoCount = 1
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 1: .ParaValue = RTrim(KHinbn)
    
        If Not .DBStored Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     構成マスタ情報の抽出に失敗しました。(構成数取得)   "
                .MB_Title = "抽出処理"
                .MB_Button = Error
                If Not .MBOX Then Exit Function
            End With
        End If
    End With
    
    gSL_Select = "SELECT 構成数 FROM 構成メンテワーク"
    gSL_Select = gSL_Select & " WHERE 子品番 = '" & RTrim(Kohinbn) & "'"
    If KOWRead(gSL_Select, 1, 2) Then
        wKoseiSu = KOW.構成数
    Else
        wKoseiSu = 1
    End If
    
    KoseiSu_Get = True
End Function
'**************************************************************
'*      「大親品番」内示テーブル - 出庫予定
'**************************************************************
Private Sub Oya_NJT_Set(NHinbn As String, Optional DBNo As Integer = 1)
    Dim wOByte      As Byte
    
    wN1DaySu = 0
    wN2DaySu = 0
    
    If wKoseiSu = 0 Then wKoseiSu = 1                       '2012/10/31
    
    '================ 2010/08/09 Start
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  jyt.納入日"
    gSL_Select = gSL_Select & ", sum(jyt.納入数) 納入数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  受注テーブル jyt"
    gSL_Select = gSL_Select & " where jyt.納入日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and jyt.品番 = '" & RTrim(NHinbn) & "'"
    gSL_Select = gSL_Select & "   and jyt.分納回数 = 0"
    gSL_Select = gSL_Select & " group by jyt.納入日"
    gSL_Select = gSL_Select & " order by jyt.納入日"
    If JYTRead(gSL_Select, , DBNo) Then
        Do Until JYT_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(JYT.納入日, "@@@@/@@/@@")) + 1
'            wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + JYT.納入数                     '2010/09/01
            wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + (JYT.納入数 * wKoseiSu)         '2010/09/01
            Call JYTReadNext(DBNo)
            
            wDSP_Flg = 1                '2010/08/23
        
        Loop
    End If
    Call JYTClose(DBNo)
    
    If wDSP_Flg = 0 Then Exit Sub                   '2010/08/23
    
    '================ 2010/08/09 End
    
    '未来在庫処理
    If H1Op_Zaiko(0) Then
        '大親品番の最終納入日を取得
    
        gSL_Select = ""
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  max(jyt.納入日) 納入日"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  受注テーブル jyt"
        gSL_Select = gSL_Select & " where jyt.納入日 between '" & wYmdS & "' and '" & wYmdE & "'"
        gSL_Select = gSL_Select & "   and jyt.品番 = '" & RTrim(NHinbn) & "'"
'        gSL_Select = gSL_Select & "   and jyt.納入先 = '" & RTrim(wTorcd) & "'"
        Call JYTRead(gSL_Select, 1, DBNo)
        
        '===================== 2010/08/23 Start
        gStr = JYT.納入日
        gSL_Select = ""
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  max(syk.出荷日) 出荷日" & vbCrLf
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  受注テーブル jyt" & vbCrLf
        gSL_Select = gSL_Select & " left outer join 出荷テーブル syk " & vbCrLf
        gSL_Select = gSL_Select & " on syk.出荷指示書NO = jyt.出荷指示書NO " & vbCrLf
        gSL_Select = gSL_Select & " and syk.出荷指示書行NO = jyt.出荷指示書行NO " & vbCrLf
        gSL_Select = gSL_Select & " where isnull(jyt.出荷実績区分,0) <> 0 " & vbCrLf
        gSL_Select = gSL_Select & "   and isnull(syk.出荷日,'') between '" & wYmdS & "' and '" & wYmdE & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and jyt.品番 = '" & RTrim(NHinbn) & "'" & vbCrLf
        Call SKTRead(gSL_Select, 1, DBNo)
        If gStr <= SKT.出荷日 Then
           gStr = SKT.出荷日
        End If
        
        gSL_Select = ""
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  max(jkt.実績日) 実績日" & vbCrLf
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  実績管理テーブル jkt" & vbCrLf
        gSL_Select = gSL_Select & "  where isnull(jkt.実績日,'') between '" & wYmdS & "' and '" & wYmdE & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and jkt.品番 = '" & RTrim(NHinbn) & "'" & vbCrLf
        Call JKTRead(gSL_Select, 1, DBNo)
        If gStr <= JKT.実績日 Then
           gStr = JKT.実績日
        End If
        
        JYT.納入日 = gStr
        gStr = ""
        '===================== 2010/08/23 End
        '===================== 2010/09/01 Start
        If JYT.納入日 < wNYMD Then
            JYT.納入日 = wNYMD
        End If
        
        wNYMD = JYT.納入日
        '===================== 2010/09/01 End
    End If


'=============== 当月(入力年月)処理
    gSL_Select = ""
    gSL_Select = gSL_Select & "select "
    For wOByte = 1 To 31
        gSL_Select = gSL_Select & " sum(内示数量" & Format(wOByte, "00") & ") 内示数量" & Format(wOByte, "00") & " , "
    Next wOByte
    gSL_Select = gSL_Select & " sum(内示数) 内示数,sum(翌月内示数) 翌月内示数,sum(翌々月内示数) 翌々月内示数,"
    gSL_Select = gSL_Select & " 内示年月,品番,納入先 "
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  内示テーブル njt"
    gSL_Select = gSL_Select & " where njt.内示年月 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   and njt.品番 = '" & RTrim(NHinbn) & "'"
    gSL_Select = gSL_Select & " group by 内示年月,品番,納入先"
    If NJTRead(gSL_Select, , DBNo) Then
        Do Until NJT_RDSTS = False
'=================== 2010/09/01 Start           '日量内示が入力されていない場合
            gLong = 0
            For wOByte = 1 To 31
                gLong = gLong + NJT.内示数量(wOByte)
            Next wOByte
            If gLong <> 0 Then
'=================== 2010/09/01 End
                '当月
                For wOByte = 1 To 31
                    If NJT.内示数量(wOByte) <> 0 Then
    '=================== 2010/08/09 Start
    '                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(iH1_YM & Format(wOByte, "00"), "@@@@/@@/@@")) + 1
    '                    wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + NJT.内示数量(wOByte)
    '=================== 2010/08/09 End
                        
                        '未来在庫処理
                        If H1Op_Zaiko(0) Then
                            wStr1 = NJT.内示年月 & Format(wOByte, "00")
                            If wStr1 > JYT.納入日 And wStr1 <= wYmdE And NJT.内示数量(wOByte) <> 0 Then
                                WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'============================ 2010/09/01 Start
'                                wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + NJT.内示数量(wOByte)
'                                wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + NJT.内示数量(wOByte)
'                                wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + NJT.内示数量(wOByte)            '2010/08/09
                                wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + (NJT.内示数量(wOByte) * wKoseiSu)
                                wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + (NJT.内示数量(wOByte) * wKoseiSu)
                                wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + (NJT.内示数量(wOByte) * wKoseiSu)
'============================ 2010/09/01 End
                                wDSP_Jyu_Nai(wY, WDd) = 1
'                            Else                                                                            '2010/09/01
'                                WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1       '2010/09/01
'                                wDSP_Jyu_Nai(wY, WDd) = 0                                                   '2010/09/01
                            End If
    '                        If wStr1 > JYT.納入日 And wStr1 <= wYmdE And Mid(wN_1KadoKB, wOByte, 1) = 0 Then
    '                            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
    '                            wDSP_Jyu_Nai(wY, WDd) = 1
    '                        End If
                        End If
                    
                    End If
                Next wOByte
'=================== 2010/09/01 Start
            Else
                If NJT.内示数 <> 0 Then
                    If wNKadosu <> 0 Then
                        wN1DaySu = NJT.内示数 / wNKadosu
                        If wN1DaySu <> 0 Then
                            For wOByte = 1 To 31
                                 If Mid(wN0KadoKB, wOByte, 1) = 0 Then
'                                     If wStr1 > JYT.納入日 And wStr1 <= wYmdE Then         '2010/08/09
'                                         WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wN1YM & Format(wOByte, "00"), "@@@@/@@/@@")) + 1
'                                         wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + wN1DaySu
'                                     End If
                                 
                                     '未来在庫処理
                                     If H1Op_Zaiko(0) Then
                                        wStr1 = NJT.内示年月 & Format(wOByte, "00")
                                         
                                         If Mid(wN0KadoKB, wOByte, 1) = 0 Then
                                            If wStr1 > JYT.納入日 And wStr1 <= wYmdE And wN1DaySu <> 0 Then
                                                WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'============================= 2010/09/01 Start
'                                                wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + wN1DaySu
'                                                wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + wN1DaySu
'                                                wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + wN1DaySu
                                                wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + (wN1DaySu * wKoseiSu)
                                                wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + (wN1DaySu * wKoseiSu)
                                                wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + (wN1DaySu * wKoseiSu)
'============================= 2010/09/01 End
                                                wDSP_Jyu_Nai(wY, WDd) = 1
                                            End If
                                         End If
                                     End If
                                     
                                 
                                 End If
                            Next wOByte
                        End If
                    End If
                End If
            End If
'=================== 2010/09/01 End
            '----- 翌月以降の内示数は稼働日数で割った日当数
            '翌月
            If NJT.翌月内示数 <> 0 Then
               If wN1Kadosu <> 0 Then
                   wN1DaySu = NJT.翌月内示数 / wN1Kadosu
                   If wN1DaySu <> 0 Then
                       For wOByte = 1 To CisFun.Val2(Mid(wYmdE, 7, 2))
                            If Mid(wN1KadoKB, wOByte, 1) = 0 Then
'================== 2010/09/01 Start
'                                If wStr1 > JYT.納入日 And wStr1 <= wYmdE Then         '2010/08/09
'                                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wN1YM & Format(wOByte, "00"), "@@@@/@@/@@")) + 1
'                                    wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + wN1DaySu
'                                End If
'================== 2010/09/01 End
                            
                                '未来在庫処理
                                If H1Op_Zaiko(0) Then
                                    With CisFun
                                        wStr = Mid(iH1_YM, 1, 6)         '>当月
                                        .Date_RsFormat = fYM
                                        .Date_Add wStr, M, 1
                                        wStr = Mid(.Date_Result, 1, 6)
                                    End With
                        
                                    wStr1 = wStr & Format(wOByte, "00")
                                    
                                    If Mid(wN1KadoKB, wOByte, 1) = 0 Then
                                        If wStr1 > JYT.納入日 And wStr1 <= wYmdE Then
                                             WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wN1YM & Format(wOByte, "00"), "@@@@/@@/@@")) + 1
'================== 2010/09/01 Start
'                                             wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + wN1DaySu
'                                             wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + wN1DaySu
'                                             wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + wN1DaySu           '2010/09/01
                                             wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + (wN1DaySu * wKoseiSu)
                                             wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + (wN1DaySu * wKoseiSu)
                                             wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + (wN1DaySu * wKoseiSu)
'================== 2010/09/01 End
                                             wDSP_Jyu_Nai(wY, WDd) = 1
                                        End If
                                    End If
                                End If
                                
                            
                            End If
                       Next wOByte
                   End If
               End If
            End If
            
            Call NJTReadNext(DBNo)
        Loop
    End If
    Call NJTClose(DBNo)

    '=============== 前月(入力年月 - 1)処理
    gSL_Select = ""
    gSL_Select = gSL_Select & "select "
    For wOByte = 1 To 31
        gSL_Select = gSL_Select & " sum(内示数量" & Format(wOByte, "00") & ") 内示数量" & Format(wOByte, "00") & " , "
    Next wOByte
    gSL_Select = gSL_Select & " sum(内示数) 内示数,sum(翌月内示数) 翌月内示数,sum(翌々月内示数) 翌々月内示数,"
    gSL_Select = gSL_Select & " 内示年月,品番,納入先 "
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  内示テーブル njt"
    gSL_Select = gSL_Select & " where njt.内示年月 = '" & RTrim(wZenYM) & "'"
    gSL_Select = gSL_Select & "   and njt.品番 = '" & RTrim(NHinbn) & "'"
    gSL_Select = gSL_Select & " group by 内示年月,品番,納入先"
    If NJTRead(gSL_Select, , DBNo) Then
        Do Until NJT_RDSTS = False
            '当月
            For wOByte = CisFun.Val2(Mid(wYmdS, 7, 2)) To 31
                If NJT.内示数量(wOByte) <> 0 Then
                    '=============================== 2010/08/09 Start
                    'WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wZenYM & Format(wOByte, "00"), "@@@@/@@/@@")) + 1
                    'wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + NJT.内示数量(wOByte)
                    If RTrim(JYT.納入日) = "" Or Mid(JYT.納入日, 1, 6) = Mid(wYmdS, 1, 6) Then
                        WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wZenYM & Format(wOByte, "00"), "@@@@/@@/@@")) + 1
'                        wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + NJT.内示数量(wOByte)               '2010/09/01
                        wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + (NJT.内示数量(wOByte) * wKoseiSu)   '2010/09/01
                    End If
                    '=============================== 2010/08/09 End
                
                    If H1Op_Zaiko(0) Then
                        If RTrim(JYT.納入日) = "" Or Mid(JYT.納入日, 1, 6) = Mid(wYmdS, 1, 6) Then
                            wStr1 = NJT.内示年月 & Format(wOByte, "00")
                            
                            If wStr1 > JYT.納入日 And wStr1 <= wYmdE And NJT.内示数量(wOByte) <> 0 Then
                                WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'======================== 2010/09/01 Start
'                                wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + NJT.内示数量(wOByte)
'                                wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + NJT.内示数量(wOByte)
                                wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + (NJT.内示数量(wOByte) * wKoseiSu)
                                wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + (NJT.内示数量(wOByte) * wKoseiSu)
'======================== 2010/09/01 End
                            End If
                            If wStr1 > JYT.納入日 And wStr1 <= wYmdE And Mid(wN_1KadoKB, wOByte, 1) = 0 Then
                                WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
                                wDSP_Jyu_Nai(wY, WDd) = 1
                            End If
                            
                        End If
                    End If
                
                End If
            Next wOByte
            Call NJTReadNext(DBNo)
        Loop
    End If
    Call NJTClose(DBNo)


'    If H1Op_Zaiko(0) Then
'        '大親品番の最終納入日を取得
'
'        gSL_Select = ""
'        gSL_Select = gSL_Select & "select"
'        gSL_Select = gSL_Select & "  max(jyt.納入日) 納入日"
'        gSL_Select = gSL_Select & " from"
'        gSL_Select = gSL_Select & "  受注テーブル jyt"
'        gSL_Select = gSL_Select & " where jyt.納入日 between '" & wYmdS & "' and '" & wYmdE & "'"
'        gSL_Select = gSL_Select & "   and jyt.品番 = '" & RTrim(wHinbn) & "'"
'        gSL_Select = gSL_Select & "   and jyt.納入先 = '" & RTrim(wTorcd) & "'"
'        Call JYTRead(gSL_Select, 1, DBNo)
'
'        '未来在庫
'        '=============== 前月処理・・・入力年月で受注データが存在しない or 最終納入日が前月
'        If RTrim(JYT.納入日) = "" Or Mid(JYT.納入日, 1, 6) = Mid(wYmdS, 1, 6) Then
'            gSL_Select = ""
'            gSL_Select = gSL_Select & "select "
'            For wInt = 1 To 31
'                gSL_Select = gSL_Select & " sum(内示数量" & Format(wInt, "00") & ") 内示数量" & Format(wInt, "00") & " , "
'            Next wInt
'            gSL_Select = gSL_Select & " sum(内示数) 内示数,sum(翌月内示数) 翌月内示数,sum(翌々月内示数) 翌々月内示数,"
'            gSL_Select = gSL_Select & " 内示年月,品番,納入先 "
'            gSL_Select = gSL_Select & " from"
'            gSL_Select = gSL_Select & "  内示テーブル njt"
'            gSL_Select = gSL_Select & " where njt.内示年月 = '" & Mid(wYmdS, 1, 6) & "'"
'            gSL_Select = gSL_Select & "   and njt.品番 = '" & RTrim(wHinbn) & "'"
'            gSL_Select = gSL_Select & "   and njt.納入先 = '" & RTrim(wTorcd) & "'"
'            gSL_Select = gSL_Select & " group by 内示年月,品番,納入先"
'            If NJTRead(gSL_Select, 1, DBNo) Then
'                For wInt = 1 To 31
'                    wStr1 = NJT.内示年月 & Format(wInt, "00")
'                    If wStr1 > JYT.納入日 And wStr1 <= wYmdE And NJT.内示数量(wInt) <> 0 Then
'                        WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'                        wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + NJT.内示数量(wInt)
'                        wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + NJT.内示数量(wInt)
'                        wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + NJT.内示数量(wInt)
'                    End If
'                    If wStr1 > JYT.納入日 And wStr1 <= wYmdE And Mid(wN_1KadoKB, wInt, 1) = 0 Then
'                        WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'                        wDSP_Jyu_Nai(wY, WDd) = 1
'                    End If
'                Next wInt
'            End If
'        End If
'
'        '=============== 当月(入力年月)処理
'        gSL_Select = ""
'        gSL_Select = gSL_Select & "select "
'        For wInt = 1 To 31
'            gSL_Select = gSL_Select & " sum(内示数量" & Format(wInt, "00") & ") 内示数量" & Format(wInt, "00") & " , "
'        Next wInt
'        gSL_Select = gSL_Select & " sum(内示数) 内示数,sum(翌月内示数) 翌月内示数,sum(翌々月内示数) 翌々月内示数,"
'        gSL_Select = gSL_Select & " 内示年月,品番,納入先 "
'        gSL_Select = gSL_Select & " from"
'        gSL_Select = gSL_Select & "  内示テーブル njt"
'        gSL_Select = gSL_Select & " where njt.内示年月 = '" & RTrim(iH1_YM) & "'"
'        gSL_Select = gSL_Select & "   and njt.品番 = '" & RTrim(wHinbn) & "'"
'        gSL_Select = gSL_Select & "   and njt.納入先 = '" & RTrim(wTorcd) & "'"
'        gSL_Select = gSL_Select & " group by 内示年月,品番,納入先"
'        If NJTRead(gSL_Select, 1, DBNo) Then
'
'            '当月
'            For wInt = 1 To 31
'                wStr1 = NJT.内示年月 & Format(wInt, "00")
'                If wStr1 > JYT.納入日 And wStr1 <= wYmdE And NJT.内示数量(wInt) <> 0 Then
'                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'                    wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + NJT.内示数量(wInt)
'                    wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + NJT.内示数量(wInt)
'                    wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + NJT.内示数量(wInt)
'                End If
'                If wStr1 > JYT.納入日 And wStr1 <= wYmdE And Mid(wN0KadoKB, wInt, 1) = 0 Then
'                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'                    wDSP_Jyu_Nai(wY, WDd) = 1
'                End If
'            Next wInt
'
'            With CisFun
'                wStr = Mid(iH1_YM, 1, 6)         '>当月
'                .Date_RsFormat = fYM
'                .Date_Add wStr, M, 1
'                wStr = Mid(.Date_Result, 1, 6)
'            End With
'
'            '----- 翌月以降の内示数は稼働日数で割った日当数
'            '翌月
'            If NJT.翌月内示数 <> 0 Then
'               If wN1Kadosu <> 0 Then
'                   wN1DaySu = NJT.翌月内示数 / wN1Kadosu
'                   If wN1DaySu <> 0 Then
'                       For wInt = 1 To 31
'                           wStr1 = wStr & Format(wInt, "00")
'                           If Mid(wN1KadoKB, wInt, 1) = 0 Then
'                               If wStr1 > JYT.納入日 And wStr1 <= wYmdE Then
'                                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wN1YM & Format(wInt, "00"), "@@@@/@@/@@")) + 1
'                                    wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + wN1DaySu
'                                    wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + wN1DaySu
'                                    wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + wN1DaySu
'                                    wDSP_Jyu_Nai(wY, WDd) = 1
'                               End If
'                           End If
'                       Next wInt
'                   End If
'               End If
'            End If
'
'        End If
'
'    End If


End Sub
'******************************************'
'*      カレンダマスタ参照
'******************************************'
Private Sub wCLM_Chk()
    wZenKadosu = 0
    wNKadosu = 0
    wN1Kadosu = 0
    wN2Kadosu = 0
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "select clm.*"
    gSL_Select = gSL_Select & "  from カレンダマスタ clm"
    gSL_Select = gSL_Select & " where clm.取引先区分 = 1"
    gSL_Select = gSL_Select & "   and clm.取引先 = ''"
    gSL_Select = gSL_Select & "   and clm.年 + clm.月"
    gSL_Select = gSL_Select & "       between '" & CisFun.Mid2(wDateS, 1, 6) & "'"
    gSL_Select = gSL_Select & "           and '" & CisFun.Mid2(wDateE, 1, 6) & "'"
    gSL_Select = gSL_Select & " order by clm.年"
    gSL_Select = gSL_Select & "        , clm.月"
    If CLMRead(gSL_Select) Then
        Do Until CLM_RDSTS = False
            '前月稼働日数
            If RTrim(CLM.年) & RTrim(CLM.月) = RTrim(wZenYM) Then
               wZenKadosu = CLM.稼動日数
               wN_1KadoKB = CLM.稼動区分
            End If
            '当月稼働日数
            If RTrim(CLM.年) & RTrim(CLM.月) = RTrim(iH1_YM) Then
               wNKadosu = CLM.稼動日数
               wN0KadoKB = CLM.稼動区分
            End If
            '翌月稼働日数
            If RTrim(CLM.年) & RTrim(CLM.月) = RTrim(wN1YM) Then
               wN1Kadosu = CLM.稼動日数
            End If
            '翌々月稼働日数
            If RTrim(CLM.年) & RTrim(CLM.月) = RTrim(wN2YM) Then
               wN2Kadosu = CLM.稼動日数
            End If
        
        '   >> 処理月の判定
            If CLM.月 = CisFun.Mid2(wDateS, 5, 2) Then
'            '   [前月]
'                wb_SDay = CisFun.Mid2(wDateS, 7, 2)
'                wb_EDay = CLM.末日
                
            Else
                If CLM.月 = CisFun.Mid2(wDateE, 5, 2) Then
                '   [翌月]
                    wN1KadoKB = CLM.稼動区分
                End If
            End If
            Call CLMReadNext
        Loop
    End If
    Call CLMClose
End Sub
'**************************************************************
'*      「完成品」受注テーブル - 出荷予定
'**************************************************************
Private Sub Kansei_JYT_Set(Optional DBNo As Integer = 1)
    gSL_Select = ""
    gSL_Select = gSL_Select & "select" & vbCrLf
    gSL_Select = gSL_Select & "  jyt.納入日" & vbCrLf
    gSL_Select = gSL_Select & ", sum(jyt.納入数) 納入数" & vbCrLf
    gSL_Select = gSL_Select & " from" & vbCrLf
    gSL_Select = gSL_Select & "  受注テーブル jyt" & vbCrLf
    gSL_Select = gSL_Select & " where jyt.納入日 between '" & wYmdS & "' and '" & wYmdE & "'" & vbCrLf
    gSL_Select = gSL_Select & "   and jyt.品番 = '" & RTrim(wHinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & "   and jyt.納入先 = '" & RTrim(wTorcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & "   and jyt.分納回数 = 0" & vbCrLf
    gSL_Select = gSL_Select & " group by jyt.納入日" & vbCrLf
    gSL_Select = gSL_Select & " order by jyt.納入日" & vbCrLf
    If JYTRead(gSL_Select, , DBNo) Then
        Do Until JYT_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(JYT.納入日, "@@@@/@@/@@")) + 1
'            wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + JYT.納入数                     '2010/09/01
            wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + (JYT.納入数 * wKoseiSu)         '2010/09/01
            Call JYTReadNext(DBNo)
            
            wDSP_Flg = 1                    '2010/08/23
        
        Loop
    End If
    Call JYTClose(DBNo)
    
    If wDSP_Flg = 0 Then Exit Sub           '2010/08/23
    
    '===================== 2010/08/23 Start
    gStr = JYT.納入日
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  max(syk.出荷日) 出荷日"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  受注テーブル jyt"
    gSL_Select = gSL_Select & " left outer join 出荷テーブル syk "
    gSL_Select = gSL_Select & " on syk.出荷指示書NO = jyt.出荷指示書NO "
    gSL_Select = gSL_Select & " and syk.出荷指示書行NO = jyt.出荷指示書行NO "
    gSL_Select = gSL_Select & " where isnull(jyt.出荷実績区分,0) <> 0 "
    gSL_Select = gSL_Select & "   and isnull(syk.出荷日,'') between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and jyt.品番 = '" & RTrim(wHinbn) & "'"
    Call SKTRead(gSL_Select, 1, DBNo)
    If gStr <= SKT.出荷日 Then
       gStr = SKT.出荷日
    End If
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  max(jkt.実績日) 実績日"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  実績管理テーブル jkt"
    gSL_Select = gSL_Select & "  where isnull(jkt.実績日,'') between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and jkt.品番 = '" & RTrim(wHinbn) & "'"
    Call JKTRead(gSL_Select, 1, DBNo)
    If gStr <= JKT.実績日 Then
       gStr = JKT.実績日
    End If
    
    JYT.納入日 = gStr
    gStr = ""
    '===================== 2010/08/23 End
    '===================== 2010/09/01 Start
    If JYT.納入日 < wNYMD Then
        JYT.納入日 = wNYMD
    End If
    
    wNYMD = JYT.納入日
    '===================== 2010/09/01 End
    
    
    If H1Op_Zaiko(0) Then
        '未来在庫
        '=============== 前月処理・・・入力年月で受注データが存在しない or 最終納入日が前月
        If RTrim(JYT.納入日) = "" Or Mid(JYT.納入日, 1, 6) = Mid(wYmdS, 1, 6) Then
            gSL_Select = ""
            gSL_Select = gSL_Select & "select "
            For wInt = 1 To 31
                gSL_Select = gSL_Select & " sum(内示数量" & Format(wInt, "00") & ") 内示数量" & Format(wInt, "00") & " , "
            Next wInt
            gSL_Select = gSL_Select & " sum(内示数) 内示数,sum(翌月内示数) 翌月内示数,sum(翌々月内示数) 翌々月内示数,"
            gSL_Select = gSL_Select & " 内示年月,品番,納入先 "
            gSL_Select = gSL_Select & " from"
            gSL_Select = gSL_Select & "  内示テーブル njt"
            gSL_Select = gSL_Select & " where njt.内示年月 = '" & Mid(wYmdS, 1, 6) & "'"
            gSL_Select = gSL_Select & "   and njt.品番 = '" & RTrim(wHinbn) & "'"
            gSL_Select = gSL_Select & "   and njt.納入先 = '" & RTrim(wTorcd) & "'"
            gSL_Select = gSL_Select & " group by 内示年月,品番,納入先"
            If NJTRead(gSL_Select, 1, DBNo) Then
                For wInt = 1 To 31
                    wStr1 = NJT.内示年月 & Format(wInt, "00")
                    If wStr1 > JYT.納入日 And wStr1 <= wYmdE And NJT.内示数量(wInt) <> 0 Then
                        WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'=========================== 2010/09/01 Start
'                        wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + NJT.内示数量(wInt)
'                        wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + NJT.内示数量(wInt)
'                        wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + NJT.内示数量(wInt)
                        wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + (NJT.内示数量(wInt) * wKoseiSu)
                        wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + (NJT.内示数量(wInt) * wKoseiSu)
                        wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + (NJT.内示数量(wInt) * wKoseiSu)
'=========================== 2010/09/01 End
                        wDSP_Jyu_Nai(wY, WDd) = 1           '2010/08/23
                    End If
'====================== 2010/08/23 Start
'                    If wStr1 > JYT.納入日 And wStr1 <= wYmdE And Mid(wN_1KadoKB, wInt, 1) = 0 Then
'                        WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'                        wDSP_Jyu_Nai(wY, WDd) = 1
'                    End If
'====================== 2010/08/23 End
                Next wInt
            End If
        End If
        
        '=============== 当月(入力年月)処理
        gSL_Select = ""
        gSL_Select = gSL_Select & "select "
        For wInt = 1 To 31
            gSL_Select = gSL_Select & " sum(内示数量" & Format(wInt, "00") & ") 内示数量" & Format(wInt, "00") & " , "
        Next wInt
        gSL_Select = gSL_Select & " sum(内示数) 内示数,sum(翌月内示数) 翌月内示数,sum(翌々月内示数) 翌々月内示数,"
        gSL_Select = gSL_Select & " 内示年月,品番,納入先 "
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  内示テーブル njt"
        gSL_Select = gSL_Select & " where njt.内示年月 = '" & RTrim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "   and njt.品番 = '" & RTrim(wHinbn) & "'"
        gSL_Select = gSL_Select & "   and njt.納入先 = '" & RTrim(wTorcd) & "'"
        gSL_Select = gSL_Select & " group by 内示年月,品番,納入先"
        If NJTRead(gSL_Select, 1, DBNo) Then
            
            '当月
            For wInt = 1 To 31
                wStr1 = NJT.内示年月 & Format(wInt, "00")
                If wStr1 > JYT.納入日 And wStr1 <= wYmdE And NJT.内示数量(wInt) <> 0 Then
                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'========================= 2010/09/01 Start
'                    wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + NJT.内示数量(wInt)
'                    wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + NJT.内示数量(wInt)
'                    wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + NJT.内示数量(wInt)
                    wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + (NJT.内示数量(wInt) * wKoseiSu)
                    wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + (NJT.内示数量(wInt) * wKoseiSu)
                    wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + (NJT.内示数量(wInt) * wKoseiSu)
'========================= 2010/09/01 End
                    wDSP_Jyu_Nai(wY, WDd) = 1           '2010/08/23
                End If
'====================== 2010/08/23 Start
'                If wStr1 > JYT.納入日 And wStr1 <= wYmdE And Mid(wN0KadoKB, wInt, 1) = 0 Then
'                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wStr1, "@@@@/@@/@@")) + 1
'                    wDSP_Jyu_Nai(wY, WDd) = 1
'                End If
'====================== 2010/08/23 End
            Next wInt
            
            With CisFun
                wStr = Mid(iH1_YM, 1, 6)         '>当月
                .Date_RsFormat = fYM
                .Date_Add wStr, M, 1
                wStr = Mid(.Date_Result, 1, 6)
            End With
            
            '----- 翌月以降の内示数は稼働日数で割った日当数
            '翌月
            If NJT.翌月内示数 <> 0 Then
               If wN1Kadosu <> 0 Then
                   wN1DaySu = NJT.翌月内示数 / wN1Kadosu
                   If wN1DaySu <> 0 Then
                       For wInt = 1 To 31
                           wStr1 = wStr & Format(wInt, "00")
                           If Mid(wN1KadoKB, wInt, 1) = 0 Then
                               If wStr1 > JYT.納入日 And wStr1 <= wYmdE Then
                                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(wN1YM & Format(wInt, "00"), "@@@@/@@/@@")) + 1
'==================== 2010/09/01 Start
'                                    wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + wN1DaySu
'                                    wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + wN1DaySu
'                                    wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + wN1DaySu
                                    wDSP_JyuSu(wY, WDd) = wDSP_JyuSu(wY, WDd) + (wN1DaySu * wKoseiSu)
                                    wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + (wN1DaySu * wKoseiSu)
                                    wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + (wN1DaySu * wKoseiSu)
'==================== 2010/09/01 End
                                    wDSP_Jyu_Nai(wY, WDd) = 1
                               End If
                           End If
                       Next wInt
                   End If
               End If
            End If
            
        End If

    End If
End Sub
'**************************************************************
'*      「完成品」出荷テーブル - 出荷実績
'**************************************************************
Private Sub Kansei_SKT_Set(Optional Torcd As String = "", Optional DBNo As Integer = 1)
    Dim sZYM        As String
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  jyt.納入先"
    gSL_Select = gSL_Select & ", skt.出荷日"
    gSL_Select = gSL_Select & ", sum(skt.出荷数) 出荷数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  出荷テーブル skt "
    gSL_Select = gSL_Select & ", 受注テーブル jyt"
    gSL_Select = gSL_Select & " where skt.出荷日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and jyt.出荷指示書NO = skt.出荷指示書NO"
    gSL_Select = gSL_Select & "   and jyt.出荷指示書行NO = skt.出荷指示書行NO"
    gSL_Select = gSL_Select & "   and jyt.品番 = '" & wHinbn & "'"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  skt.出荷日"
    gSL_Select = gSL_Select & ", jyt.納入先"
    If JYTRead(gSL_Select, , DBNo) Then
        Do Until JYT_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(JYT.出荷日, "@@@@/@@/@@")) + 1
            wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + JYT.出荷数
'            If NSK.レベル = 0 And Mid(JYT.出荷日, 1, 6) = RTrim(iH1_YM) Then
            If NSK.レベル = 0 And Mid(JYT.出荷日, 1, 6) = RTrim(iH1_YM) And JYT.納入先 = wTorcd Then
               wSyuSu = wSyuSu + JYT.出荷数
            End If
            If Torcd = "" Or _
               Torcd = JYT.納入先 Then
                wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + JYT.出荷数
            End If
            Call JYTReadNext(DBNo)
        Loop
    End If
    Call JYTClose(DBNo)


    If iH1_YM <> Format(Now, "YYYYMM") Then
        With CisFun
            .Date_RsFormat = fYM
            .Date_Add RTrim(iH1_YM) & "01", M, -1
            sZYM = Mid(.Date_Result, 1, 6)
        End With
    
        gSL_Select = ""
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  jyt.納入先"
        gSL_Select = gSL_Select & ", skt.出荷日"
        gSL_Select = gSL_Select & ", sum(skt.出荷数) 出荷数"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  出荷テーブル skt "
        gSL_Select = gSL_Select & ", 受注テーブル jyt"
        gSL_Select = gSL_Select & " where substring(skt.出荷日,1,6) = '" & sZYM & "'"
        gSL_Select = gSL_Select & "   and jyt.出荷指示書NO = skt.出荷指示書NO"
        gSL_Select = gSL_Select & "   and jyt.出荷指示書行NO = skt.出荷指示書行NO"
        gSL_Select = gSL_Select & "   and jyt.品番 = '" & wHinbn & "'"
        gSL_Select = gSL_Select & "   and jyt.納入先 = '" & wTorcd & "'"
        gSL_Select = gSL_Select & " group by"
        gSL_Select = gSL_Select & "  skt.出荷日"
        gSL_Select = gSL_Select & ", jyt.納入先"
        If JYTRead(gSL_Select, , DBNo) Then
            Do Until JYT_RDSTS = False
               wZenSyuSu = wZenSyuSu + JYT.出荷数
               Call JYTReadNext(DBNo)
            Loop
        End If
        Call JYTClose(DBNo)
    End If

End Sub
'**************************************************************
'*      「完成品」発注テーブル - 入荷予定
'**************************************************************
Private Sub Kansei_HCT_Set(Optional DBNo As Integer = 1)
    Call ItemsClearHCT
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  hct.納入日"
    gSL_Select = gSL_Select & ", sum(hct.発注数) 発注数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  発注テーブル hct"
    gSL_Select = gSL_Select & " where hct.納入日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and hct.品番 = '" & RTrim(wHinbn) & "'"
    gSL_Select = gSL_Select & "   and hct.分納区分 = 0"
    gSL_Select = gSL_Select & " group by hct.納入日"
    If HCTRead(gSL_Select, , DBNo) Then
        Do Until HCT_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(HCT.納入日, "@@@@/@@/@@")) + 1
            wDSP_HacSu(wY, WDd) = wDSP_HacSu(wY, WDd) + HCT.発注数
            Call HCTReadNext(DBNo)
        Loop
    End If
    Call HCTClose(DBNo)

    
    Call ItemsClearSGT
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & " max(sgt.作業日) 作業日"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  作業日報テーブル sgt"
    gSL_Select = gSL_Select & " where sgt.作業日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and sgt.品番 = '" & RTrim(wHinbn) & "'"
    If wSyoriDspKB = 3 Or wSyoriDspKB = 5 Then
       gSL_Select = gSL_Select & "   and sgt.部門 = '" & RTrim(NSK.仕入先) & "'"
       gSL_Select = gSL_Select & "   and sgt.工順 = " & RTrim(NSK.作業順)
    End If
    Call SGTRead(gSL_Select, 1, DBNo)
       
    
    If HCT.納入日 < SGT.作業日 Then
       HCT.納入日 = SGT.作業日
    End If
    
    
    If H1Op_Zaiko(0) Then
        '未来在庫処理
        If RTrim(HCT.納入日) = "" Or Mid(HCT.納入日, 1, 6) = Mid(wYmdS, 1, 6) Then
            gSL_Select = ""
            gSL_Select = gSL_Select & "select"
            gSL_Select = gSL_Select & "  hnj.納期"
            gSL_Select = gSL_Select & ", sum(hnj.数量) 数量"
            gSL_Select = gSL_Select & " from"
            gSL_Select = gSL_Select & "  発注内示テーブル hnj"
            gSL_Select = gSL_Select & " where hnj.内示年月 = '" & Mid(wYmdS, 1, 6) & "'"
            gSL_Select = gSL_Select & "   and hnj.品番 = '" & RTrim(wHinbn) & "'"
            gSL_Select = gSL_Select & "   and hnj.納期 between '" & wYmdS & "' and '" & Mid(wYmdS, 1, 6) & "31" & "'"
            gSL_Select = gSL_Select & " group by"
            gSL_Select = gSL_Select & "  hnj.納期"
'================== 2010/09/01 Start
'            If HNJRead(gSL_Select, , DBNo) Then
'                Do Until HNJ_RDSTS = False
'                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(HNJ.納期, "@@@@/@@/@@")) + 1
'                    wDSP_HacSu(wY, WDd) = wDSP_HacSu(wY, WDd) + HNJ.数量
'                    wCMP_NyuSu(wY, WDd) = wCMP_NyuSu(wY, WDd) + HNJ.数量
'                    wDSP_NyuSu(wY, WDd) = wDSP_NyuSu(wY, WDd) + HNJ.数量
'                    wDSP_Hac_Nai(wY, WDd) = 1
'
'                    Call HNJReadNext(DBNo)
'                Loop
'            End If
'            Call HNJClose(DBNo)
'================== 2010/09/01 End
        End If
        
        If HCT.納入日 <> "" Then
            With CisFun
                .Date_RsFormat = fYMD
                .Date_Add HCT.納入日, D, 1
                wStr = .Date_Result
            End With
        Else
            wStr = RTrim(iH1_YM) & "01"
        End If
        
        
        gSL_Select = ""
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  hnj.納期"
        gSL_Select = gSL_Select & ", sum(hnj.数量) 数量"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  発注内示テーブル hnj"
        gSL_Select = gSL_Select & " where hnj.内示年月 = '" & RTrim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "   and hnj.品番 = '" & RTrim(wHinbn) & "'"
        gSL_Select = gSL_Select & "   and hnj.納期 between '" & wStr & "' and '" & wYmdE & "'"
        gSL_Select = gSL_Select & " group by"
        gSL_Select = gSL_Select & "  hnj.納期"
'============== 2010/09/01 Start
'        If HNJRead(gSL_Select, , DBNo) Then
'            Do Until HNJ_RDSTS = False
'                WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(HNJ.納期, "@@@@/@@/@@")) + 1
'                wDSP_HacSu(wY, WDd) = wDSP_HacSu(wY, WDd) + HNJ.数量
'                wCMP_NyuSu(wY, WDd) = wCMP_NyuSu(wY, WDd) + HNJ.数量
'                wDSP_NyuSu(wY, WDd) = wDSP_NyuSu(wY, WDd) + HNJ.数量
'                wDSP_Hac_Nai(wY, WDd) = 1
'
'                Call HNJReadNext(DBNo)
'            Loop
'        End If
'        Call HNJClose(DBNo)
'============== 2010/09/01 End

    End If

'    gSL_Select = ""
'    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  sis.指示日"
'    gSL_Select = gSL_Select & ", sum(sis.指示数) 指示数"
'    gSL_Select = gSL_Select & " from"
'    gSL_Select = gSL_Select & "  生産指示テーブル sis"
'    gSL_Select = gSL_Select & " where sis.指示日 between '" & wYmdS & "' and '" & wYmdE & "'"
'    gSL_Select = gSL_Select & "   and sis.品番 = '" & RTrim(wHinbn) & "'"
'    gSL_Select = gSL_Select & "   and sis.部署CD = '" & RTrim(wTorcd) & "'"
'    gSL_Select = gSL_Select & "   and sis.データ区分 = 'K'"
'    gSL_Select = gSL_Select & "   and sis.工順 = ( select max(sis2.工順) from 生産指示テーブル sis2"
'    gSL_Select = gSL_Select & "                           where sis2.指示日 between '" & wYmdS & "' and '" & wYmdE & "'"
'    gSL_Select = gSL_Select & "                           and   sis2.品番 = '" & RTrim(wHinbn) & "'"
'    gSL_Select = gSL_Select & "                           and   sis2.部署CD = '" & RTrim(wTorcd) & "'"
'    gSL_Select = gSL_Select & "                           and   sis2.データ区分 = 'K' ) "
'    gSL_Select = gSL_Select & " group by sis.指示日"
'    If SSTRead(gSL_Select, , DBNo) Then
'        Do Until SST_RDSTS = False
'            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(SST.指示日, "@@@@/@@/@@")) + 1
'            wDSP_HacSu(wY, WDd) = wDSP_HacSu(wY, WDd) + SST.指示数
'            Call SSTReadNext(DBNo)
'        Loop
'    End If
'    Call SSTClose(DBNo)

End Sub
'**************************************************************
'*      入荷テーブル - 入荷実績
'**************************************************************
Private Sub Kyotu_NYT_Set(Optional Torcd As String = "", Optional DBNo As Integer = 1)
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  nyt.仕入先"
    gSL_Select = gSL_Select & ", nyt.入荷日"
    gSL_Select = gSL_Select & ", sum(nyt.入荷数) 入荷数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  入荷テーブル nyt"
    gSL_Select = gSL_Select & " where nyt.入荷日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and nyt.品番 = '" & RTrim(wHinbn) & "'"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  nyt.入荷日"
    gSL_Select = gSL_Select & ", nyt.仕入先"
    If NYTRead(gSL_Select, , DBNo) Then
        Do Until NYT_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(NYT.入荷日, "@@@@/@@/@@")) + 1
            wCMP_NyuSu(wY, WDd) = wCMP_NyuSu(wY, WDd) + NYT.入荷数
            If Torcd = "" Or _
               Torcd = NYT.仕入先 Then
                wDSP_NyuSu(wY, WDd) = wDSP_NyuSu(wY, WDd) + NYT.入荷数
            End If
            Call NYTReadNext(DBNo)
        Loop
    End If
    Call NYTClose(DBNo)


    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  sgt.部門"
    gSL_Select = gSL_Select & ", sgt.作業日"
    gSL_Select = gSL_Select & ", sum(sgt.加工数) 加工数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  作業日報テーブル sgt"
    gSL_Select = gSL_Select & " where sgt.作業日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and sgt.品番 = '" & RTrim(wHinbn) & "'"
    If wSyoriDspKB = 3 Or wSyoriDspKB = 5 Then
       gSL_Select = gSL_Select & "   and sgt.部門 = '" & RTrim(NSK.仕入先) & "'"
       gSL_Select = gSL_Select & "   and sgt.工順 = " & RTrim(NSK.作業順)
    End If
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  sgt.作業日"
    gSL_Select = gSL_Select & ", sgt.部門"
    If SGTRead(gSL_Select, , DBNo) Then
        Do Until SGT_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(SGT.作業日, "@@@@/@@/@@")) + 1
            wCMP_NyuSu(wY, WDd) = wCMP_NyuSu(wY, WDd) + SGT.加工数
            If Torcd = "" Or _
               Torcd = SGT.部門 Then
                wDSP_NyuSu(wY, WDd) = wDSP_NyuSu(wY, WDd) + SGT.加工数
            End If
            Call SGTReadNext(DBNo)
        Loop
    End If
    Call SGTClose(DBNo)

End Sub
'**************************************************************
'*      「部品構成品」実績管理テーブル - 出荷実績＆入荷実績
'**************************************************************
Private Sub Buhin_JKT_Set(Optional Kubun As Integer = 0, Optional DBNo As Integer = 1)
    Dim sZYM        As String
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  jkt.実績日"
    gSL_Select = gSL_Select & ", jkt.実績区分"
    gSL_Select = gSL_Select & ", sum(jkt.実績数) 実績数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  実績管理テーブル jkt"
    gSL_Select = gSL_Select & " where jkt.実績日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and jkt.品番 = '" & RTrim(wHinbn) & "'"
    If wSyoriDspKB = 3 Or wSyoriDspKB = 5 Then
       gSL_Select = gSL_Select & "   and jkt.実績先 = '" & RTrim(NSK.仕入先) & "'"
       gSL_Select = gSL_Select & "   and jkt.工順 = " & NSK.作業順
    End If
    gSL_Select = gSL_Select & "   and jkt.実績区分 in (1, 2, 3, 4)"
    gSL_Select = gSL_Select & "   and isnull(jkt.親品番,0) = 0 "
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  jkt.実績日"
    gSL_Select = gSL_Select & ", jkt.実績区分"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  jkt.実績日"
    gSL_Select = gSL_Select & ", jkt.実績区分"
    If JKTRead(gSL_Select, , DBNo) Then
        Do Until JKT_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(JKT.実績日, "@@@@/@@/@@")) + 1
            Select Case JKT.実績区分
                Case 1
                    If Kubun = 0 Or _
                       Kubun = 1 Then
                        wDSP_NyuSu(wY, WDd) = wDSP_NyuSu(wY, WDd) + JKT.実績数
                        wCMP_NyuSu(wY, WDd) = wCMP_NyuSu(wY, WDd) + JKT.実績数
                    End If
                Case 2, 3, 4
                    If Kubun = 0 Or _
                       Kubun = 2 Then
                       wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + JKT.実績数
                       wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + JKT.実績数
                    
                       If Kubun = 2 Then
                          If NSK.レベル = 0 And Mid(JKT.実績日, 1, 6) = RTrim(iH1_YM) Then
                             wSyuSu = wSyuSu + JKT.実績数
                          End If
                       End If
                    End If
            End Select
            Call JKTReadNext(DBNo)
        Loop
    End If
    Call JKTClose(DBNo)


    If iH1_YM <> Format(Now, "YYYYMM") And Kubun = 2 Then
        With CisFun
            .Date_Add RTrim(iH1_YM) & "01", M, -1
            sZYM = Mid(.Date_Result, 1, 6)
        End With
    
        gSL_Select = ""
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  jkt.実績日"
        gSL_Select = gSL_Select & ", jkt.実績区分"
        gSL_Select = gSL_Select & ", sum(jkt.実績数) 実績数"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  実績管理テーブル jkt"
        gSL_Select = gSL_Select & " where substring(jkt.実績日,1,6) = '" & sZYM & "'"
        gSL_Select = gSL_Select & "   and jkt.品番 = '" & RTrim(wHinbn) & "'"
        gSL_Select = gSL_Select & "   and jkt.実績区分 in (2, 3)"
        gSL_Select = gSL_Select & "   and isnull(jkt.親品番,0) = 0 "
        gSL_Select = gSL_Select & " group by"
        gSL_Select = gSL_Select & "  jkt.実績日"
        gSL_Select = gSL_Select & ", jkt.実績区分"
        gSL_Select = gSL_Select & " order by"
        gSL_Select = gSL_Select & "  jkt.実績日"
        gSL_Select = gSL_Select & ", jkt.実績区分"
      
        If JKTRead(gSL_Select, , DBNo) Then
            Do Until JKT_RDSTS = False
               wZenSyuSu = wZenSyuSu + JKT.実績数
               Call JKTReadNext(DBNo)
            Loop
        End If
        Call JKTClose(DBNo)
    
    End If

'    '大親品番の最終実績日を取得
'
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "select"
'    gSL_Select = gSL_Select & "  max(skt.出荷日) 納入日"
'    gSL_Select = gSL_Select & " from"
'    gSL_Select = gSL_Select & "  出荷テーブル skt "
'    gSL_Select = gSL_Select & ", 受注テーブル jyt"
'    gSL_Select = gSL_Select & " where skt.出荷日 between '" & wYmdS & "' and '" & wYmdE & "'"
'    gSL_Select = gSL_Select & "   and jyt.出荷指示書NO = skt.出荷指示書NO"
'    gSL_Select = gSL_Select & "   and jyt.出荷指示書行NO = skt.出荷指示書行NO"
'    gSL_Select = gSL_Select & "   and jyt.品番 = '" & RTrim(wHinbn) & "'"
'    gSL_Select = gSL_Select & "   and jyt.納入先 = '" & RTrim(wTorcd) & "'"
'    Call JYTRead(gSL_Select, 1, DBNo)
    

End Sub
'**************************************************************
'*      「部品構成品」発注内示テーブル - 入荷予定
'**************************************************************
Private Sub Buhin_HNJ_Set(Optional Torcd As String = "", Optional DBNo As Integer = 1)
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  hnj.納期"
    gSL_Select = gSL_Select & ", sum(hnj.数量) 数量"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  発注内示テーブル hnj"
    gSL_Select = gSL_Select & " where hnj.内示年月 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   and hnj.品番 = '" & RTrim(wHinbn) & "'"
    If Torcd <> "" Then
        gSL_Select = gSL_Select & "   and hnj.手配先 = '" & Torcd & "'"
    End If
    gSL_Select = gSL_Select & "   and hnj.納期 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  hnj.納期"
    If HNJRead(gSL_Select, , DBNo) Then
        Do Until HNJ_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(HNJ.納期, "@@@@/@@/@@")) + 1
            wDSP_HacSu(wY, WDd) = wDSP_HacSu(wY, WDd) + HNJ.数量
            Call HNJReadNext(DBNo)
        Loop
    End If
    Call HNJClose(DBNo)
End Sub
'**************************************************************
'*      「材料構成品」材料入荷テーブル - 入荷実績
'**************************************************************
Private Sub Zairyo_ZYT_Set(Optional DBNo As Integer = 1)
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  zyt.入荷日"
    gSL_Select = gSL_Select & ", sum(zyt.入荷数) 入荷数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料入荷テーブル zyt"
    gSL_Select = gSL_Select & " where zyt.入荷日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and zyt.材料管理番号 = " & wZKanriNo
    gSL_Select = gSL_Select & " group by zyt.入荷日"
    gSL_Select = gSL_Select & " order by zyt.入荷日"
    If ZYTRead(gSL_Select, , DBNo) Then
        Do Until ZYT_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(ZYT.入荷日, "@@@@/@@/@@")) + 1
            wDSP_NyuSu(wY, WDd) = wDSP_NyuSu(wY, WDd) + ZYT.入荷数
            wCMP_NyuSu(wY, WDd) = wCMP_NyuSu(wY, WDd) + ZYT.入荷数
            Call ZYTReadNext(DBNo)
        Loop
    End If
    Call ZYTClose(DBNo)
End Sub
'**************************************************************
'*      「材料構成品」材料実績管理テーブル - 出荷実績＆入荷実績
'**************************************************************
Private Sub Zairyo_ZJK_Set(Optional DBNo As Integer = 1)
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  zjk.実績日"
    gSL_Select = gSL_Select & ", zjk.実績区分"
    gSL_Select = gSL_Select & ", sum(zjk.実績数) 実績数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料実績管理テーブル zjk"
    gSL_Select = gSL_Select & " where zjk.実績日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and 材質 = '" & RTrim(wZaisitu) & "'"
    gSL_Select = gSL_Select & "   and 板厚 = " & wItaatu
    gSL_Select = gSL_Select & "   and 幅 = " & wHaba
    gSL_Select = gSL_Select & "   and 長さ = " & wNagasa
    gSL_Select = gSL_Select & "   and 実績区分 in (1, 2, 3)"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  zjk.実績日"
    gSL_Select = gSL_Select & ", zjk.実績区分"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  zjk.実績日"
    gSL_Select = gSL_Select & ", zjk.実績区分"
    If ZJKRead(gSL_Select, , DBNo) Then
        Do Until ZJK_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(ZJK.実績日, "@@@@/@@/@@")) + 1
            Select Case ZJK.実績区分
                Case 1
                    wDSP_NyuSu(wY, WDd) = wDSP_NyuSu(wY, WDd) + ZJK.実績数
                    wCMP_NyuSu(wY, WDd) = wCMP_NyuSu(wY, WDd) + ZJK.実績数
                Case 2, 3
                    wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + ZJK.実績数
                    wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + ZJK.実績数
            End Select
            Call ZJKReadNext(DBNo)
        Loop
    End If
    Call ZJKClose(DBNo)

    

    '材料使用量テーブル(出荷実績)
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  zsr.作業日"
    gSL_Select = gSL_Select & ", sum(zsr.使用量) 使用量"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料使用量テーブル zsr"
    gSL_Select = gSL_Select & " where zsr.作業日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and 材質 = '" & RTrim(wZaisitu) & "'"
    gSL_Select = gSL_Select & "   and 板厚 = " & wItaatu
    gSL_Select = gSL_Select & "   and 幅 = " & wHaba
    gSL_Select = gSL_Select & "   and 長さ = " & wNagasa
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  zsr.作業日"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  zsr.作業日"
    If ZSRRead(gSL_Select, , DBNo) Then
        Do Until ZSR_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(ZSR.作業日, "@@@@/@@/@@")) + 1
        
            wDSP_SyuSu(wY, WDd) = wDSP_SyuSu(wY, WDd) + ZSR.使用量
            wCMP_SyuSu(wY, WDd) = wCMP_SyuSu(wY, WDd) + ZSR.使用量
            
            Call ZSRReadNext(DBNo)
        Loop
    End If
    
    
End Sub
'**************************************************************
'*      「材料構成品」材料発注内示テーブル - 入荷予定
'**************************************************************
Private Sub Zairyo_ZNJ_Set(Optional DBNo As Integer = 1)
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  znj.納期"
    gSL_Select = gSL_Select & ", sum(znj.数量) 数量"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料発注内示テーブル znj"
    gSL_Select = gSL_Select & " where znj.内示年月 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   and znj.材料管理番号 = " & wZKanriNo
    gSL_Select = gSL_Select & "   and znj.納期 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  znj.納期"
    If ZNJRead(gSL_Select, , DBNo) Then
        Do Until ZNJ_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(ZNJ.納期, "@@@@/@@/@@")) + 1
            wDSP_HacSu(wY, WDd) = wDSP_HacSu(wY, WDd) + ZNJ.数量
            Call ZNJReadNext(DBNo)
        Loop
    End If
    Call ZNJClose(DBNo)
End Sub
'**************************************************************
'*      「材料構成品」材料発注テーブル - 入荷予定
'**************************************************************
Private Sub Zairyo_ZCT_Set(Optional DBNo As Integer = 1)
    Call ItemsClearZCT
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  zct.納入日"
    gSL_Select = gSL_Select & ", sum(zct.発注数) 発注数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料発注テーブル zct"
    gSL_Select = gSL_Select & " where zct.納入日 between '" & wYmdS & "' and '" & wYmdE & "'"
    gSL_Select = gSL_Select & "   and zct.材料管理番号 = " & wZKanriNo
    gSL_Select = gSL_Select & "   and zct.分納区分 = 0"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  zct.納入日"
    If ZCTRead(gSL_Select, , DBNo) Then
        Do Until ZCT_RDSTS = False
            WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(ZCT.納入日, "@@@@/@@/@@")) + 1
            wDSP_HacSu(wY, WDd) = wDSP_HacSu(wY, WDd) + ZCT.発注数
            Call ZCTReadNext(DBNo)
        Loop
    End If
    Call ZCTClose(DBNo)

    If H1Op_Zaiko(0) Then
        '未来在庫
        If RTrim(ZCT.納入日) = "" Or Mid(ZCT.納入日, 1, 6) = Mid(wYmdS, 1, 6) Then
            gSL_Select = ""
            gSL_Select = gSL_Select & "select"
            gSL_Select = gSL_Select & "  znj.納期"
            gSL_Select = gSL_Select & ", sum(znj.数量) 数量"
            gSL_Select = gSL_Select & " from"
            gSL_Select = gSL_Select & "  材料発注内示テーブル znj"
            gSL_Select = gSL_Select & " where znj.内示年月 = '" & Mid(wYmdS, 1, 6) & "'"
            gSL_Select = gSL_Select & "   and znj.材料管理番号 = " & wZKanriNo
            gSL_Select = gSL_Select & "   and znj.納期 between '" & wYmdS & "' and '" & wYmdE & "'"
            gSL_Select = gSL_Select & " group by"
            gSL_Select = gSL_Select & "  znj.納期"
'================== 2010/09/01 Start
'            If ZNJRead(gSL_Select, , DBNo) Then
'                Do Until ZNJ_RDSTS = False
'                    WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(ZNJ.納期, "@@@@/@@/@@")) + 1
'                    wDSP_HacSu(wY, WDd) = wDSP_HacSu(wY, WDd) + ZNJ.数量
'                    wCMP_NyuSu(wY, WDd) = wCMP_NyuSu(wY, WDd) + ZNJ.数量
'                    wDSP_NyuSu(wY, WDd) = wDSP_NyuSu(wY, WDd) + ZNJ.数量
'                    wDSP_Hac_Nai(wY, WDd) = 1
'
'                    Call ZNJReadNext(DBNo)
'                Loop
'            End If
'            Call ZNJClose(DBNo)
'================== 2010/09/01 End
        End If
        
        If ZCT.納入日 <> "" Then
            With CisFun
                .Date_RsFormat = fYMD
                .Date_Add ZCT.納入日, D, 1
                wStr = .Date_Result
            End With
        Else
            wStr = RTrim(iH1_YM) & "01"
        End If
        
        gSL_Select = ""
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  znj.納期"
        gSL_Select = gSL_Select & ", sum(znj.数量) 数量"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  材料発注内示テーブル znj"
        gSL_Select = gSL_Select & " where znj.内示年月 = '" & RTrim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "   and znj.材料管理番号 = " & wZKanriNo
        gSL_Select = gSL_Select & "   and znj.納期 between '" & wStr & "' and '" & wYmdE & "'"
        gSL_Select = gSL_Select & " group by"
        gSL_Select = gSL_Select & "  znj.納期"
'================= 2010/09/01 Start
'        If ZNJRead(gSL_Select, , DBNo) Then
'            Do Until ZNJ_RDSTS = False
'                WDd = DateDiff("d", Format(wYmdS, "@@@@/@@/@@"), Format(ZNJ.納期, "@@@@/@@/@@")) + 1
'                wDSP_HacSu(wY, WDd) = wDSP_HacSu(wY, WDd) + ZNJ.数量
'                wCMP_NyuSu(wY, WDd) = wCMP_NyuSu(wY, WDd) + ZNJ.数量
'                wDSP_NyuSu(wY, WDd) = wDSP_NyuSu(wY, WDd) + ZNJ.数量
'                wDSP_Hac_Nai(wY, WDd) = 1
'
'                Call ZNJReadNext(DBNo)
'            Loop
'        End If
'================= 2010/09/01 End
        Call ZNJClose(DBNo)
    End If

End Sub

'+==============================================================================+
'+
'+          グリッド処理
'+
'+==============================================================================+
'
'+----------------------------------------------------------+
'+      グリッド    《 Scroll  イベント   》
'+----------------------------------------------------------+
Private Sub vsGrid2_Scroll()
    vsGrid1.TopRow = vsGrid2.TopRow
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 RowColChange  イベント   》
'+----------------------------------------------------------+
Private Sub vsGrid2_RowColChange()
    If ProcHB = "B1" Then
        wApl_RowNo = vsGrid2.Row
        wApl_ColNo = vsGrid2.Col
        Call EditMenu(wApl_RowNo)
    End If
End Sub
'+==============================================================================+
'+
'+          セル内のメニュー処理
'+
'+==============================================================================+
'+---------------------------+
'+      ファンクション押下
'+---------------------------+
Private Function Push_PFKey(Key_Code As Integer, Shift As Integer)
    Dim ws_PFKey    As String
    Dim wi_Menu     As Integer

    Push_PFKey = False
    Select Case Key_Code
           Case vbKeyF1: ws_PFKey = "F1"
           Case vbKeyF2: ws_PFKey = "F2"
           Case vbKeyF3: ws_PFKey = "F3"
           Case vbKeyF4: ws_PFKey = "F4"
           Case vbKeyF5: ws_PFKey = "F5"
           Case vbKeyF6: ws_PFKey = "F6"
           Case vbKeyF7: ws_PFKey = "F7"
           Case vbKeyF8: ws_PFKey = "F8"
           Case Else:    Exit Function
    End Select

    For wi_Menu = 1 To 8
        If CisFun.Mid2(PB_Apl(wi_Menu).Caption, 1, 2) = ws_PFKey Then
            Call PB_Apl_Click(wi_Menu)
            Exit For
        End If
    Next wi_Menu

    Push_PFKey = True
End Function
'+---------------------------+
'+      メニュークリック
'+---------------------------+
Private Sub PB_Apl_Click(Index As Integer)
    If PB_Apl(Index).Tag <> "" Then
        Call Load_Program(Index)
    '   ::_____<< 下記の処理は、ﾌｧﾝｸｼｮﾝ･ｷｰを続けて押下した時の対応 >>_____::
        ProcHB = "B2"
        DoEvents
        ProcHB = "B1"
    '   ::^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    End If
End Sub
'+---------------------------+
'+      プログラム起動
'+---------------------------+
Private Sub Load_Program(Index As Integer)
    Dim wl_ProcID   As Long
    Dim wi_Kind     As Integer
    Dim wi_Menu     As Integer
    Dim ws_Apl      As String
    Dim ws_SYmd     As String
    Dim ws_NaiYM    As String       '>内示年月
    Dim ws_TaiYM    As String       '>対象年月
    Dim ws_Hinbn    As String
    Dim ws_Torcd    As String
    Dim ws_Ukeir    As String
    Dim ws_Zaist    As String       '>材質
    Dim wd_Itatu    As Double       ' 板厚
    Dim wd_Width    As Double       ' 幅
    Dim wd_Long     As Double       ' 長さ
    Dim wb_JKbn     As Byte         '>実績区分

'   >> 各値の設定
    With CisFun
        .Date_RsFormat = fYMD
        .Date_Add wDateS, D, wApl_ColNo - 3
        ws_SYmd = .Date_Result
    End With
    ws_NaiYM = RTrim(iH1_YM)
    ws_TaiYM = wDSP_YM
    ws_Hinbn = RTrim(vsGrid2.TextMatrix(wApl_RowNo, 54))
    gInt = (wApl_RowNo - 1) \ 5
    ws_Torcd = cTorcd(gInt)
    ws_Ukeir = cUkeir(gInt)
    ws_Zaist = RTrim(vsGrid2.TextMatrix(wApl_RowNo, 57))
    wd_Itatu = CisFun.Val2(vsGrid2.TextMatrix(wApl_RowNo, 58))
    wd_Width = CisFun.Val2(vsGrid2.TextMatrix(wApl_RowNo, 59))
    wd_Long = CisFun.Val2(vsGrid2.TextMatrix(wApl_RowNo, 60))
    wi_Kind = CisFun.Val2(PB_Apl(Index).Tag)
    Select Case wApl_RowNo Mod 5
        Case 1, 2: wb_JKbn = 2
        Case 3, 4: wb_JKbn = 1
    End Select
    wi_Menu = Index
'   >> 起動判定
    If Right(wApl_EXE(wi_Kind, wi_Menu), 4) = ".EXE" Then
    '   >> ﾌﾟﾛｸﾞﾗﾑ名のｾｯﾄ
        ws_Apl = gIniExe & wApl_EXE(wi_Kind, wi_Menu)
    '   >> ﾊﾟﾗﾒｰﾀのｾｯﾄ
        Select Case wApl_PAM(wi_Kind, wi_Menu)
            Case "年月,品番"
                ws_Apl = ws_Apl & " " & CisFun.Mid2(ws_SYmd, 1, 6) _
                                & "," & ws_Hinbn
            Case "年月,品番,取引先,受入"
                ws_Apl = ws_Apl & " " & CisFun.Mid2(ws_SYmd, 1, 6) _
                                & "," & ws_Hinbn _
                                & "," & ws_Torcd _
                                & "," & ws_Ukeir
            Case "年月x2,品番,取引先,受入"
                ws_Apl = ws_Apl & " " & ws_NaiYM _
                                & "," & ws_TaiYM _
                                & "," & ws_Hinbn _
                                & "," & ws_Torcd _
                                & "," & ws_Ukeir & ",SetOnly"
            Case "年月,材質,サイズ"
                ws_Apl = ws_Apl & " " & CisFun.Mid2(ws_SYmd, 1, 6) _
                                & "," & ws_Zaist _
                                & "," & wd_Itatu _
                                & "," & wd_Width _
                                & "," & wd_Long
            Case "年月,材質,サイズ,取引先,受入"
                ws_Apl = ws_Apl & " " & CisFun.Mid2(ws_SYmd, 1, 6) _
                                & "," & ws_Zaist _
                                & "," & wd_Itatu _
                                & "," & wd_Width _
                                & "," & wd_Long _
                                & "," & ws_Torcd _
                                & "," & ws_Ukeir
            Case "年月x2,材質,サイズ,取引先,受入"
                ws_Apl = ws_Apl & " " & ws_NaiYM _
                                & "," & ws_TaiYM _
                                & "," & ws_Zaist _
                                & "," & wd_Itatu _
                                & "," & wd_Width _
                                & "," & wd_Long _
                                & "," & ws_Torcd _
                                & "," & ws_Ukeir & ",SetOnly"
            Case "取引先,受入,日付"
                ws_Apl = ws_Apl & " " & ws_Torcd _
                                & "," & ws_Ukeir _
                                & "," & ws_SYmd
            Case "実績区分"
                ws_Apl = ws_Apl & " " & wb_JKbn
        End Select
    '   >> ﾌﾟﾛｸﾞﾗﾑ起動
        wl_ProcID = Shell(ws_Apl, vbNormalFocus)
    Else
    '   <><><> 組込み <><><>
        Select Case wApl_EXE(wi_Kind, wi_Menu)
            Case "CXJ0410": Call Load_CXJ0410(ws_SYmd, iH1_Hinbn, iH1_Torcd, "")
            Case "CXH0410": Call Load_CXH0410(ws_SYmd, ws_Hinbn, ws_Torcd, "")
            Case "CXHZ0410": Call Load_CXHZ0410(ws_SYmd, ws_Zaist, wd_Itatu, wd_Width, wd_Long)
        End Select
    End If

End Sub
'+---------------------------+
'+      受注出荷状況
'+---------------------------+
Private Sub Load_CXJ0410(SYmd As String, Hinbn As String, Torcd As String, Ukeir As String)
    If vsGrid2.TextMatrix(wApl_RowNo, wApl_ColNo) = "" Then
        RV_SYmd = CisFun.Mid2(SYmd, 1, 6)
    Else
        RV_SYmd = RTrim(SYmd)
    End If
    RV_Hinbn = RTrim(Hinbn)
    RV_Seban = ""
    RV_Torcd = RTrim(Torcd)
    RV_Ukeir = RTrim(Ukeir)
    Rv_Call3 = "K"
    Select Case wApl_RowNo
        Case 1: Rv_CallRtn = 0
        Case 2: Rv_CallRtn = 1
    End Select
    Rv_Call3Rtn = 1
'
    'Set Rv_CallObj = CSA0040
    Set Rv_CallObj = CSA0070
    Set Rv_CallObj2 = CXJ0425
'
    CXJ0425.Show
    'CSA0040.Hide
    CSA0070.Hide
    CXJ0410.Show
End Sub
'+---------------------------+
'+      発注入荷状況
'+---------------------------+
Private Sub Load_CXH0410(SYmd As String, Hinbn As String, Torcd As String, Ukeir As String)
    If vsGrid2.TextMatrix(wApl_RowNo, wApl_ColNo) = "" Then
        RV_SYmd = CisFun.Mid2(SYmd, 1, 6)
    Else
        RV_SYmd = RTrim(SYmd)
    End If
    RV_Hinbn = RTrim(Hinbn)
    RV_Seban = ""

'    Select Case wApl_RowNo
    Select Case (wApl_RowNo Mod c_MRow)
'
        Case 3:     RV_Torcd = "":           RV_Ukeir = "":           Rv_CallRtn = 0
        Case 4:     RV_Torcd = "":           RV_Ukeir = "":           Rv_CallRtn = 1
        Case Else:  RV_Torcd = RTrim(Torcd): RV_Ukeir = RTrim(Ukeir): Rv_CallRtn = 1
    End Select
    Rv_Call3 = "K"
'
    'Set Rv_CallObj = CSA0040
    Set Rv_CallObj = CSA0070
    Set Rv_CallObj2 = CXH0425
'
    CXH0425.Show
    'CSA0040.Hide
    CSA0070.Hide
    CXH0410.Show
End Sub
'+---------------------------+
'+      材料発注入荷状況
'+---------------------------+
Private Sub Load_CXHZ0410(SYmd As String, Zaist As String, Itatu As Double, Haba As Double, Nagasa As Double)
    If vsGrid2.TextMatrix(wApl_RowNo, wApl_ColNo) = "" Then
        RV_SYmd = CisFun.Mid2(SYmd, 1, 6)
    Else
        RV_SYmd = RTrim(SYmd)
    End If
    RV_Zaist = RTrim(Zaist)
    RV_Seban = ""
    RV_Torcd = ""
    RV_Ukeir = ""
    RV_Itatu = Itatu
    RV_Width = Haba
    RV_Long = Nagasa
    Rv_Call3 = "K"

'
'    Rv_CallRtn = 1
    Select Case (wApl_RowNo Mod c_MRow)
        Case 4:     Rv_CallRtn = 1
        Case Else:  Rv_CallRtn = 0
    End Select
'
    'Set Rv_CallObj = CSA0040
    Set Rv_CallObj = CSA0070
    Set Rv_CallObj2 = CXHZ0425
'
    CXHZ0425.Show
    'CSA0040.Hide
    CSA0070.Hide
    CXHZ0410.Show
End Sub
'
'+--------------------------------+
'+      メニュー内容の読込
'+--------------------------------+
Private Function ReadMenu(Optional DBNo As Integer = 0) As Boolean
    Dim wi_Kind     As Integer
    Dim wi_Menu     As Integer

    ReadMenu = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  sym.*"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  名称マスタ sym"
    gSL_Select = gSL_Select & " where sym.区分名称 like '納入指示%'"
    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
    If SYMRead(gSL_Select, , DBNo) Then
        Do Until SYM_RDSTS = False
            Select Case SYM.区分名称
                Case "納入指示（親）出荷予定":   wi_Kind = 1
                Case "納入指示（親）出荷実績":   wi_Kind = 2
                Case "納入指示（親）入荷予定":   wi_Kind = 3
                Case "納入指示（親）入荷実績":   wi_Kind = 4
                Case "納入指示（親）在庫":       wi_Kind = 5
                Case "納入指示（部品）出荷予定": wi_Kind = 6
                Case "納入指示（部品）出荷実績": wi_Kind = 7
                Case "納入指示（部品）入荷予定": wi_Kind = 8
                Case "納入指示（部品）入荷実績": wi_Kind = 9
                Case "納入指示（部品）在庫":     wi_Kind = 10
                Case "納入指示（材料）出荷予定": wi_Kind = 11
                Case "納入指示（材料）出荷実績": wi_Kind = 12
                Case "納入指示（材料）入荷予定": wi_Kind = 13
                Case "納入指示（材料）入荷実績": wi_Kind = 14
                Case "納入指示（材料）在庫":     wi_Kind = 15
            End Select
            wi_Menu = SYM.数字1
            wApl_TTL(wi_Kind, wi_Menu) = SYM.値名称
            wApl_EXE(wi_Kind, wi_Menu) = SYM.英数字1
            wApl_PAM(wi_Kind, wi_Menu) = SYM.英数字2
            wApl_CTL(wi_Kind, wi_Menu) = SYM.英数字3
            Call SYMReadNext(DBNo)
        Loop
    End If
    Call SYMClose(DBNo)

    ReadMenu = True
End Function
'+--------------------------------+
'+      メニュー内容の編集
'+--------------------------------+
Private Function EditMenu(RowNo As Long) As Boolean
    Dim wi_Kind     As Integer
    Dim wi_Menu     As Integer
    Dim wb_Edit     As Boolean

    EditMenu = False

    Select Case RowNo
        Case 1 To 5: wi_Kind = RowNo
        Case Else
            Select Case vsGrid2.TextMatrix(RowNo, 56)
                Case "B": wi_Kind = (RowNo - 1) Mod 5 + 6
                Case "Z": wi_Kind = (RowNo - 1) Mod 5 + 11
            End Select
    End Select

    For wi_Menu = 1 To 8
        If wApl_EXE(wi_Kind, wi_Menu) = "" Then
            wb_Edit = False
        Else
            wb_Edit = True
        End If
        Select Case wApl_CTL(wi_Kind, wi_Menu)
            Case "社内": If vsGrid2.TextMatrix(RowNo, 61) = "0" Then wb_Edit = False
            Case "社外": If vsGrid2.TextMatrix(RowNo, 61) = "1" Then wb_Edit = False
        End Select
        If wb_Edit Then
            PB_Apl(wi_Menu).Caption = wApl_TTL(wi_Kind, wi_Menu)
            PB_Apl(wi_Menu).Tag = wi_Kind
            PB_Apl(wi_Menu).Enabled = True
        Else
            PB_Apl(wi_Menu).Caption = ""
            PB_Apl(wi_Menu).Tag = ""
            PB_Apl(wi_Menu).Enabled = False
        End If
    Next wi_Menu
'   >> 表示制御

    EditMenu = True
End Function
'+--------------------------------+
'+      名称マスタ初期設定
'+--------------------------------+
Private Sub MeisyoInit()
    Call ItemsClearSYMCreate
    With SYMCreate
        .制御区分 = "X"
        .作成者 = gTanto
'   >> 納入指示（親）出荷予定
        .区分名称 = "納入指示（親）出荷予定"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "受注画面入力"
        .英数字1 = "PXJ0010.EXE"
        .英数字2 = ""
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
            .値 = "5"
            .値名称 = "受注出荷状況照会"
            .英数字1 = "CXJ0410"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 5
            Call SYMInitCreate
            .値 = "6"
            .値名称 = "受注出荷データ修正"
            .英数字1 = "PXJ0910.EXE"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 6
            Call SYMInitCreate
        End If
'   >> 納入指示（親）出荷実績
        .区分名称 = "納入指示（親）出荷実績"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "出荷実績入力"
        .英数字1 = "PXJ0210.EXE"
        .英数字2 = ""
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
            .値 = "5"
            .値名称 = "受注出荷状況照会"
            .英数字1 = "CXJ0410"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 5
            Call SYMInitCreate
            .値 = "6"
            .値名称 = "受注出荷データ修正"
            .英数字1 = "PXJ0910.EXE"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 6
            Call SYMInitCreate
        End If
'   >> 納入指示（親）入荷予定
        .区分名称 = "納入指示（親）入荷予定"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "発注入力【単品】"
        .英数字1 = "PXH0010.EXE"
        .英数字2 = "取引先,受入,日付"
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
            .値 = "2"
            .値名称 = "発注入力【カレンダ】"
            .英数字1 = "PXH0020.EXE"
            .英数字2 = "年月,品番,取引先,受入"
            .英数字3 = ""
            .数字1 = 2
            Call SYMInitCreate
            .値 = "5"
            .値名称 = "発注入荷状況照会"
            .英数字1 = "CXH0410"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 5
            Call SYMInitCreate
            .値 = "6"
            .値名称 = "発注入荷データ修正"
            .英数字1 = "PXH0910.EXE"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 6
            Call SYMInitCreate
        End If
'   >> 納入指示（親）入荷実績
        .区分名称 = "納入指示（親）入荷実績"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "入荷実績入力【一覧】"
        .英数字1 = "PXH0210.EXE"
        .英数字2 = ""
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
            .値 = "3"
            .値名称 = "入荷画面入力"
            .英数字1 = "PXH0230.EXE"
            .英数字2 = "取引先,受入,日付"
            .英数字3 = ""
            .数字1 = 3
            Call SYMInitCreate
'            .値 = "4"
'            .値名称 = "入出庫実績入力"
'            .英数字1 = "PXH0500.EXE"
'            .英数字2 = "実績区分"
'            .英数字3 = ""
'            .数字1 = 4
            .値 = "4"
            .値名称 = "作業実績入力"
            .英数字1 = "PSJ0010.EXE"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 4
            Call SYMInitCreate
            .値 = "5"
            .値名称 = "発注入荷状況照会"
            .英数字1 = "CXH0410"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 5
            Call SYMInitCreate
            .値 = "6"
            .値名称 = "発注入荷データ修正"
            .英数字1 = "PXH0910.EXE"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 6
            Call SYMInitCreate
        End If
'   >> 納入指示（親）在庫
        .区分名称 = "納入指示（親）在庫"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "在庫メンテ"
        .英数字1 = "PXZ0010.EXE"
        .英数字2 = "年月,品番"
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
        End If

'   >> 納入指示（部品）出荷実績
'        .区分名称 = "納入指示（部品）出荷実績"
'        .区分桁数 = 1
'        .値 = "4"
'        .値名称 = "入出庫実績入力"
'        .英数字1 = "PXH0500.EXE"
'        .英数字2 = "実績区分"
'        .英数字3 = ""
'        .数字1 = 4
'        If Not SYMInitCreate(True) Then
'        End If
'   >> 納入指示（部品）入荷予定
        .区分名称 = "納入指示（部品）入荷予定"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "発注内示入力"
        .英数字1 = "PNJ0030.EXE"
        .英数字2 = "年月x2,品番,取引先,受入"
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
        End If
'   >> 納入指示（部品）入荷実績
        .区分名称 = "納入指示（部品）入荷実績"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "入荷実績入力【一覧】"
        .英数字1 = "PXH0210.EXE"
        .英数字2 = ""
        .英数字3 = "社外"
        .数字1 = 1
        If Not SYMInitCreate(True) Then
            .値 = "2"
'            .値名称 = "作業実績入力"
'            .英数字1 = "PSA0030.EXE"
'            .英数字2 = ""
'            .英数字3 = "社内"
'            .数字1 = 2
            .値 = "3"
            .値名称 = "入荷画面入力"
            .英数字1 = "PXH0230.EXE"
            .英数字2 = "取引先,受入,日付"
            .英数字3 = ""
            .数字1 = 3
            Call SYMInitCreate
'            .値 = "4"
'            .値名称 = "入出庫実績入力"
'            .英数字1 = "PXH0500.EXE"
'            .英数字2 = "実績区分"
'            .英数字3 = ""
'            .数字1 = 4
'            Call SYMInitCreate
            .値名称 = "作業実績入力"
            .英数字1 = "PSJ0010.EXE"
            .英数字2 = ""
            .英数字3 = "社内"
            .数字1 = 4
            Call SYMInitCreate
            .値 = "5"
            .値名称 = "発注入荷状況照会"
            .英数字1 = "CXH0410"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 5
            Call SYMInitCreate
            .値 = "6"
            .値名称 = "発注入荷データ修正"
            .英数字1 = "PXH0910.EXE"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 6
            Call SYMInitCreate
        End If
'   >> 納入指示（部品）在庫
        .区分名称 = "納入指示（部品）在庫"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "在庫メンテ"
        .英数字1 = "PXZ0010.EXE"
        .英数字2 = "年月,品番"
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
        End If

'   >> 納入指示（材料）出荷実績
'        .区分名称 = "納入指示（材料）出荷実績"
'        .区分桁数 = 1
'        .値 = "4"
'        .値名称 = "材料入出庫実績入力"
'        .英数字1 = "PXHZ0500.EXE"
'        .英数字2 = "実績区分"
'        .英数字3 = ""
'        .数字1 = 4
        .区分名称 = "納入指示（材料）出荷実績"
        .区分桁数 = 1
        .値 = "4"
        .値名称 = "材料使用量入力"
        .英数字1 = "PSJ0020.EXE"
        .英数字2 = "実績区分"
        .英数字3 = ""
        .数字1 = 4
        If Not SYMInitCreate(True) Then
        End If
'   >> 納入指示（材料）入荷予定
        .区分名称 = "納入指示（材料）入荷予定"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "材料発注内示入力"
        .英数字1 = "PNJZ0030.EXE"
        .英数字2 = "年月x2,材質,サイズ,取引先,受入"
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
        End If
'   >> 納入指示（材料）入荷実績
        .区分名称 = "納入指示（材料）入荷実績"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "材料入荷実績入力"
        .英数字1 = "PXHZ0210.EXE"
        .英数字2 = ""
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
            .値 = "3"
            .値名称 = "材料入荷画面入力"
            .英数字1 = "PXHZ0230.EXE"
            .英数字2 = "取引先,受入,日付"
            .英数字3 = ""
            .数字1 = 3
            Call SYMInitCreate
'            .値 = "4"
'            .値名称 = "材料入出庫実績入力"
'            .英数字1 = "PXHZ0500.EXE"
'            .英数字2 = "実績区分"
'            .英数字3 = ""
'            .数字1 = 4
            .値 = "4"
            .値名称 = "材料使用量入力"
            .英数字1 = "PSJ0020.EXE"
            .英数字2 = "実績区分"
            .英数字3 = ""
            .数字1 = 4
            Call SYMInitCreate
            .値 = "5"
            .値名称 = "材料発注入荷状況照会"
            .英数字1 = "CXHZ0410"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 5
            Call SYMInitCreate
            .値 = "6"
            .値名称 = "材料発注入荷データ修正"
            .英数字1 = "PXHZ0910.EXE"
            .英数字2 = ""
            .英数字3 = ""
            .数字1 = 6
            Call SYMInitCreate
        End If
'   >> 納入指示（材料）在庫
        .区分名称 = "納入指示（材料）在庫"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "材料在庫メンテ"
        .英数字1 = "PXZ0020.EXE"
        .英数字2 = "年月,材質,サイズ"
        .英数字3 = ""
        .数字1 = 1
        If Not SYMInitCreate(True) Then
        End If
    End With
End Sub
'--------------------------------------------------------------------- 2010/06/03 Insert Start
'+--------------------------------+
'+      納入指示進捗データ作成
'+--------------------------------+
Private Sub NouData_Edit()
    'NouData_Edit = False
    
    With CisDB
        .SQL = "納入指示進捗データ作成_取引先"
        .StoadoCount = 15
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用             '処理区分
        .ParaNo = 2: .ParaIO = Input用             '取引先
        .ParaNo = 3: .ParaIO = Input用             '入力年月
        .ParaNo = 4: .ParaIO = Input用             '取引先区分
        .ParaNo = 5: .ParaIO = Input用             '未来在庫計算区分
        .ParaNo = 6: .ParaIO = OutPut用            '年月日開始
        .ParaNo = 7: .ParaIO = OutPut用            '年月日終了
        .ParaNo = 8: .ParaIO = OutPut用            '抽出用年月日開始
        .ParaNo = 9: .ParaIO = OutPut用            '抽出用年月日終了
        .ParaNo = 10: .ParaIO = OutPut用            '抽出用年月日終了
        .ParaNo = 11: .ParaIO = OutPut用            '抽出用年月日終了
        .ParaNo = 12: .ParaIO = OutPut用            '抽出用年月日終了
        .ParaNo = 13: .ParaIO = OutPut用            'エラー番号(OUTPUT)
        .ParaNo = 14: .ParaIO = OutPut用            'エラー出力件数
        .ParaNo = 15: .ParaIO = OutPut用            '正常出力件数

        .ParaNo = 1: .ParaValue = wSyoriDspKB
        .ParaNo = 2: .ParaValue = iH1_Torcd
        .ParaNo = 3: .ParaValue = iH1_YM
        If H1Op_Torkb(0).Value Then
            .ParaNo = 4: .ParaValue = 0
        Else
            .ParaNo = 4: .ParaValue = 1
        End If
        If H1Op_Zaiko(0).Value Then
            .ParaNo = 5: .ParaValue = 0
        Else
            .ParaNo = 5: .ParaValue = 1
        End If
'
        If Not .DBStored Then
            MsgBox "ストアドエラー "
 '           GoTo Update_Proc_Err
            Exit Sub
        End If

        .ParaNo = 13
        If .ParaValue <> 0 Then Exit Sub

        .ParaNo = 6: wDateS = .ParaValue
        .ParaNo = 7: wDateE = .ParaValue
        .ParaNo = 8: wYmdS = .ParaValue
        .ParaNo = 9: wYmdE = .ParaValue
        .ParaNo = 10: wZenNaiSu = .ParaValue
        .ParaNo = 11: wZenSyuSu = .ParaValue
        .ParaNo = 12: wSyuSu = .ParaValue
    End With
    
    Call NouData_Set
        
    'NouData_Edit = True
': Update_Proc_Err

End Sub
'+--------------------------------+
'+      データを配列に格納
'+--------------------------------+
Private Sub NouData_Set()
    Erase wDSP_JyuSu
    Erase wDSP_HacSu
    Erase wDSP_SyuSu
    Erase wDSP_NyuSu
    Erase wCMP_ZaiSu
    Erase wCMP_ZaiKb
    Erase wCMP_NyuSu
    Erase wCMP_SyuSu
    Erase cHinbn
    Erase cZKanriNo
    Erase cTorcd
    Erase cUkeir
    Erase cKoujyun
    
    Erase wDSP_Jyu_Nai
    Erase wDSP_Hac_Nai

    'NouData_Set = False
    
    Key_Hinban = ""
    Key_Torcd = ""
    Key_Ukeir = ""            ' 2010/07/27 Insert
    wY = 0
    wKensu = 0
    With CisDB
'----------------------------------------------------------------------------------------------- 2010/07/27 Update Start
'        gSL_Select = "SELECT NSK.順位,NSK.品番,NSK.大親納入先,NSK.部品材料区分, ISNULL(NSS.区分,0) 区分"
        gSL_Select = "SELECT NSK.順位,NSK.品番,NSK.大親納入先,NSK.受入,NSK.部品材料区分, ISNULL(NSS.区分,0) 区分"
'----------------------------------------------------------------------------------------------- 2010/07/27 Update End
        For wInt = 1 To 72
            gSL_Select = gSL_Select & "   , SUM(CONVERT(INT,ISNULL(NSS.値" + CStr(wInt) + ",'0'))) 値" + CStr(wInt)
        Next wInt
        gSL_Select = gSL_Select & "  FROM 納入指示構成ワーク NSK"
        gSL_Select = gSL_Select & "  LEFT OUTER JOIN 納入指示進捗ワーク NSS"
        gSL_Select = gSL_Select & "    ON NSS.処理端末 = NSK.処理端末"
        gSL_Select = gSL_Select & "   AND NSS.品番 = NSK.大親品番"
'        gSL_Select = gSL_Select & "   AND NSS.取引先 = NSK.大親納入先"
        gSL_Select = gSL_Select & " WHERE NSK.処理端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & "   AND NSK.展開区分 in (0, 1)"
'----------------------------------------------------------------------------------------------- 2010/07/27 Update Start
'       gSL_Select = gSL_Select & " GROUP BY NSK.順位,NSK.品番,NSK.大親納入先,NSK.部品材料区分,NSS.区分"
        gSL_Select = gSL_Select & " GROUP BY NSK.順位,NSK.品番,NSK.大親納入先,NSK.受入,NSK.部品材料区分,NSS.区分"
'----------------------------------------------------------------------------------------------- 2010/07/27 Update End
        gSL_Select = gSL_Select & " ORDER BY NSK.順位, NSK.品番, NSK.大親納入先, NSK.受入, NSK.部品材料区分, NSS.区分"
        
        If NSKRead(gSL_Select, , 0) Then
            Key_Hinban = NSK.品番
            Key_Torcd = NSK.大親納入先
            Key_Ukeir = NSK.受入            ' 2010/07/27 Insert
            Do Until Not NSK_RDSTS
                With NSK
                    For wInt = 1 To 72
                        If .値(wInt) <> "0" Then
                            Select Case .区分
'============================================================== 2010/07/20 Start
'                                Case 1: wDSP_JyuSu(wY, wInt) = CInt(.値(wInt))
'                                Case 2: wDSP_HacSu(wY, wInt) = CInt(.値(wInt))
'                                Case 3: wDSP_SyuSu(wY, wInt) = CInt(.値(wInt))
'                                Case 4: wDSP_NyuSu(wY, wInt) = CInt(.値(wInt))
'                                Case 5: wCMP_ZaiSu(wY, wInt) = CInt(.値(wInt))
'                                Case 7: wCMP_NyuSu(wY, wInt) = CInt(.値(wInt))
'                                Case 8: wCMP_SyuSu(wY, wInt) = CInt(.値(wInt))
                                
                                Case 1: wDSP_JyuSu(wY, wInt) = CLng(.値(wInt))
                                Case 2: wDSP_HacSu(wY, wInt) = CLng(.値(wInt))
                                Case 3: wDSP_SyuSu(wY, wInt) = CLng(.値(wInt))
                                Case 4: wDSP_NyuSu(wY, wInt) = CLng(.値(wInt))
                                Case 5: wCMP_ZaiSu(wY, wInt) = CLng(.値(wInt))
                                Case 7: wCMP_NyuSu(wY, wInt) = CLng(.値(wInt))
                                Case 8: wCMP_SyuSu(wY, wInt) = CLng(.値(wInt))
'============================================================== 2010/07/20 End
                                
                                Case 6: wCMP_ZaiKb(wY, wInt) = Trim(.値(wInt))
                                Case 9: wDSP_Jyu_Nai(wY, wInt) = CByte(.値(wInt))
                                Case 10: wDSP_Hac_Nai(wY, wInt) = CByte(.値(wInt))
                            End Select
                        End If
                    Next wInt
                End With
                
                Call NSKReadNext(0)
'------------------------------------------------------------------------------ 2010/07/27 Update Start
'                If Key_Hinban <> NSK.品番 Or Key_Torcd <> NSK.大親納入先 Or Not NSK_RDSTS Then
                If Key_Hinban <> NSK.品番 Or Key_Torcd <> NSK.大親納入先 Or Key_Ukeir <> NSK.受入 Or Not NSK_RDSTS Then
'------------------------------------------------------------------------------ 2010/07/27 Update End
                    cHinbn(wY) = RTrim(Key_Hinban)
                    wY = wY + 1
                    wKensu = wKensu + 1
                    Key_Hinban = NSK.品番
                    Key_Torcd = NSK.大親納入先
                    Key_Ukeir = NSK.受入            ' 2010/07/27 Insert
                End If
            Loop
        End If
    End With
        
   'NouData_Set = True
End Sub
'--------------------------------------------------------------------- 2010/06/03 Insert End


