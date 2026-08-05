VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CNJ3010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "内示一覧入力"
   ClientHeight    =   11100
   ClientLeft      =   15
   ClientTop       =   960
   ClientWidth     =   15315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11100
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin VB.PictureBox H0_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   5610
      ScaleHeight     =   675
      ScaleWidth      =   3255
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   540
      Width           =   3315
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   615
         Left            =   30
         Top             =   30
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   1085
         ForeColor       =   16711680
         Caption         =   "表示順"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.OptionButton H0Op_Sort 
            BackColor       =   &H00C0C0C0&
            Caption         =   "納入先"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2130
            TabIndex        =   12
            Top             =   300
            Width           =   1005
         End
         Begin VB.OptionButton H0Op_Sort 
            BackColor       =   &H00C0C0C0&
            Caption         =   "背番号"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1050
            TabIndex        =   11
            Top             =   300
            Width           =   1005
         End
         Begin VB.OptionButton H0Op_Sort 
            BackColor       =   &H00C0C0C0&
            Caption         =   "品番"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   10
            Top             =   300
            Width           =   765
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   8190
      Left            =   60
      ScaleHeight     =   8130
      ScaleWidth      =   15015
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2640
      Width           =   15075
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   8070
         Left            =   30
         Top             =   15
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   14235
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   7995
            Left            =   30
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   30
            Width           =   14895
            _Version        =   196608
            _ExtentX        =   26273
            _ExtentY        =   14102
            _StockProps     =   228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "　　|品番　　　　　　　　|納入先　　　　　　　　　　　|受入|背番号|収容数|［前月］当月内示|当月内示　|翌月内示　|翌々月内示|"
            Rows            =   24
            Cols            =   20
            BackColor       =   12648384
            BackColorFixed  =   14737632
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            GridLines       =   2
            SelectionMode   =   1
            Begin CisText_V60.CisText iB1_Nai1 
               Height          =   375
               Left            =   10440
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   330
               Visible         =   0   'False
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   661
               cFormat         =   "##,###,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##,###,###"
               MaxLength       =   10
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "########"
               cILength        =   8
            End
            Begin CisText_V60.CisText iB1_Nai2 
               Height          =   375
               Left            =   11745
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   330
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               cFormat         =   "##,###,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##,###,###"
               MaxLength       =   10
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "########"
               cILength        =   8
            End
            Begin CisText_V60.CisText iB1_Nai3 
               Height          =   375
               Left            =   13035
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   330
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               cFormat         =   "##,###,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##,###,###"
               MaxLength       =   10
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "########"
               cILength        =   8
            End
            Begin Cis3D_v60.CIS3D PNL_Wait 
               Height          =   930
               Left            =   4725
               Top             =   2700
               Visible         =   0   'False
               Width           =   5160
               _ExtentX        =   9102
               _ExtentY        =   1640
               BackColor       =   16761024
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderWidth     =   5
               Begin Cis3D_v60.CIS3D Message 
                  Height          =   525
                  Left            =   240
                  Top             =   210
                  Width           =   4680
                  _ExtentX        =   8255
                  _ExtentY        =   926
                  BackColor       =   16761024
                  ForeColor       =   0
                  Caption         =   "【抽出中】しばらくお待ち下さい"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
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
                  cBoderWidth     =   2
                  cAlingnment     =   7
                  cBoderStyle     =   1
               End
            End
         End
      End
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
      Height          =   1245
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   14925
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1335
      Width           =   14985
      Begin Cis3D_v60.CIS3D Data_Back1 
         Height          =   1080
         Left            =   11115
         Top             =   60
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   1905
         BackColor       =   12632319
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
         Begin Cis3D_v60.CIS3D Data_Back 
            Height          =   975
            Left            =   60
            Top             =   60
            Width           =   3630
            _ExtentX        =   6403
            _ExtentY        =   1720
            ForeColor       =   16711680
            Caption         =   "【未登録ﾃﾞｰﾀ】参照元"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   60
            Begin VB.OptionButton H1Op_Data 
               BackColor       =   &H00C0FFC0&
               Caption         =   "全て"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   0
               Left            =   105
               TabIndex        =   7
               Top             =   495
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton H1Op_Data 
               BackColor       =   &H00C0C0C0&
               Caption         =   "前月内示"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   1
               Left            =   1140
               TabIndex        =   8
               Top             =   495
               Width           =   1275
            End
            Begin VB.OptionButton H1Op_Data 
               BackColor       =   &H00C0C0C0&
               Caption         =   "品番ﾏｽﾀ"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   2
               Left            =   2400
               TabIndex        =   9
               Top             =   495
               Width           =   1215
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1080
         Left            =   8655
         Top             =   60
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1905
         ForeColor       =   16711680
         Caption         =   ""
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
         cAlingnment     =   6
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   1050
            Left            =   15
            Top             =   15
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   1852
            ForeColor       =   16711680
            Caption         =   "表示内容"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   2
            cPositionX      =   15
            cPositionY      =   15
         End
         Begin VB.OptionButton H1Op_Dsp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "登録済ﾃﾞｰﾀのみ"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   495
            TabIndex        =   5
            Top             =   375
            Width           =   1875
         End
         Begin VB.OptionButton H1Op_Dsp 
            BackColor       =   &H00C0FFC0&
            Caption         =   "未登録ﾃﾞｰﾀのみ"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   495
            TabIndex        =   4
            Top             =   30
            Value           =   -1  'True
            Width           =   1875
         End
         Begin VB.OptionButton H1Op_Dsp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "全て"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   495
            TabIndex        =   6
            Top             =   720
            Width           =   900
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1080
         Index           =   2
         Left            =   1665
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1905
         ForeColor       =   255
         Caption         =   "契約先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   2
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
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
            Height          =   360
            Left            =   1080
            Top             =   480
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1080
         Index           =   5
         Left            =   60
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   1905
         ForeColor       =   255
         Caption         =   "内示年月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   60
         Begin CisYM_V60.CisYM iH1_Ym 
            Height          =   360
            Left            =   90
            TabIndex        =   1
            Top             =   480
            Width           =   1395
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
            cChkResult      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1080
         Index           =   0
         Left            =   5340
         Top             =   60
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1905
         ForeColor       =   16711680
         Caption         =   "品番"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   75
            TabIndex        =   3
            Top             =   480
            Width           =   3195
            _ExtentX        =   5636
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
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
      Height          =   720
      Left            =   9000
      ScaleHeight     =   660
      ScaleWidth      =   6045
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   555
      Width           =   6105
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3480
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   4320
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   5160
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   1755
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   900
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   11
         BtnCaption      =   "切替"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711935
         BtnShadowColor  =   16777215
         PFCaption       =   "( F6 )"
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
         PFBackColor     =   16711935
      End
      Begin CisBtn_60.CisBtn PB_DSP 
         Height          =   585
         Left            =   45
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   11
         BtnCaption      =   "表示"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   0
         BtnShadowColor  =   16777215
         PFCaption       =   "( F5 )"
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
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   2625
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   11
         BtnCaption      =   "削除"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   255
         BtnShadowColor  =   16777215
         PFCaption       =   "( F3 )"
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
         PFBackColor     =   255
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【内示一覧入力】"
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
         Left            =   14385
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
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
         Height          =   300
         Left            =   11835
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   0
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
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   345
         Left            =   9885
         Top             =   15
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   609
         Caption         =   "表示件数"
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
         cAlingnment     =   1
         cPositionX      =   50
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   285
            Left            =   990
            Top             =   30
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            BackColor       =   16777152
            Caption         =   "##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
      End
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   15
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      End
   End
   Begin VB.Timer Timer2 
      Left            =   705
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   -150
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   -15
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
End
Attribute VB_Name = "CNJ3010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  内示一覧入力
'**       フォームID    :  CNJ3010
'**       処理概要      :
'**
'**       作  成  日    :  2006/10/13
'**       変  更  日    :  2009/03/12 By CIS 横山興業用カスタマイズ
'**       変  更  日    :  2010/01/26 By CIS 条件(品番)追加
'**       変　更　日    :  2010/06/29 By CIS 土岐津展開,行色変え
'**       変　更　日    :  2010/07/07 By CIS 背番号幅増
'**       変　更　日    :  2010/09/01 By CIS 内示年月取得変更
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim Up_Flg                  As Byte
    Dim CisVsGrid1              As New CisVsGrid3
    Dim mColor_On               As Long
    Dim mColor_Off              As Long
'   グリッドの背景色
    Dim mColorOrg               As Long
    Dim mColorChg               As Long
    Dim mColorDel               As Long
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT START
    Dim mColorSel               As Long
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT END
    Dim wInt                    As Integer
'   行
    Dim NowRow                  As Long
    Dim NewRow                  As Long
'   [前月]表示切替ﾌﾗｸﾞ
    Dim Chg_Flg                 As Byte
'   ｸﾞﾘｯﾄﾞ表示切替
    Dim Dsp_Flg                 As Byte
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》            +
'+----------------------------------------------------------+
Private Sub Form_Load()
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

'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 12
        .NumberSet = NoSpace
        .EditCheckCol = 12
        .FixedHeight = 330
        .PatanMax = 4
        .SelectCol = 20
                
        .VSSort
        
        .InitGet = False
        .Init
    End With
    Call GridInit1

    Chg_Flg = 0: Dsp_Flg = 0
    
'年月初期表示
    Call GetNJProc_YM(3)
    iH1_Ym = gn_HNYM

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = ""
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    '===================== 2010/09/01 Start
    gSL_Select = "SELECT MAX(内示年月) 内示年月 FROM 内示テーブル "
    gSL_Select = gSL_Select & " Where SUBSTRING(内示年月,1,2) = '20'"
    If NJTRead(gSL_Select, 1) Then
       iH1_Ym = NJT.内示年月
    End If
    '===================== 2010/09/01 End

'   背景色の設定
    mColor_On = H1Op_Dsp(1).BackColor
    mColor_Off = H1Op_Dsp(0).BackColor
'   背景色の設定（グリッド）
    mColorOrg = vsGrid1.BackColor
    mColorChg = &HC0E0FF
    mColorDel = &H8080FF
    
    H1Op_Dsp(1).Value = True
    H1Op_Data(0).Value = True
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT START
'    mColorSel = &HFFFFC0               '2010/06/28
    mColorSel = &H80FFFF                '2010/06/28
    H0Op_Sort(0).Value = True
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT END
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False

    Select Case Key_Code
        Case vbKeyF5:       If PB_DSP.Visible Then Call PB_DSP_Click        '【画面】
        Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click        '【切替】
        Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click        '【削除】
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeySpace:    If ProcHB = "B1" And SyoriKB = "DEL" Then Call VSGrid1_Click
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
                                Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                Call PB_ENT_Click
                                Dummy.Enabled = False
                            End If
        Case Else:          Exit Function
    End Select

    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp
    End
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
    If ProcHB = "B1" Then
       ProcHB = "H1"
       Call DispChange(ProcHB)
       Call HeadBodyClear("B")
       vsGrid1.HighLight = flexHighlightNever
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
        PB_DEL.BtnCaption = "戻る"
    Else
        SyoriKB = ""
        PB_DEL.BtnCaption = "削除"
    End If
    Call SyoriName(SyoriKB)
    iH1_Ym.SetFocus
    
    '削除の場合は強制的に【登録済ﾃﾞｰﾀのみ】
    If SyoriKB = "DEL" Then
       H1Op_Dsp(0).Value = False: H1Op_Dsp(1).Value = False: H1Op_Dsp(2).Value = True
       H1Op_Dsp(0).Enabled = False: H1Op_Dsp(1).Enabled = False
    Else
       H1Op_Dsp(0).Value = True: H1Op_Dsp(1).Value = False: H1Op_Dsp(2).Value = False
       H1Op_Dsp(0).Enabled = True: H1Op_Dsp(1).Enabled = True
    End If
    
End Sub
'+----------------------------+
'+    切　替  キ ー(F6)       +
'+----------------------------+
Private Sub PB_CHG_Click()
    If Chg_Flg = 0 Then
       Chg_Flg = 1
    Else
       Chg_Flg = 0
    End If

    'ｸﾞﾘｯﾄﾞ表示切替
    If Dsp_Flg = 0 Then
       If Chg_Flg = 0 Then
          Call CisVsGrid1.LayOut(0)
       Else
          Call CisVsGrid1.LayOut(1)
       End If
    Else
       If Chg_Flg = 0 Then
          Call CisVsGrid1.LayOut(2)
       Else
          Call CisVsGrid1.LayOut(3)
       End If
    End If

End Sub
'+----------------------------+
'+    表　示 キ ー(F5)       +
'+----------------------------+
Private Sub PB_DSP_Click()
    If Dsp_Flg = 0 Then
       Dsp_Flg = 1
    Else
       Dsp_Flg = 0
    End If

    'ｸﾞﾘｯﾄﾞ表示切替
    If Dsp_Flg = 0 Then
       Call CisVsGrid1.LayOut(0)
    Else
       Call CisVsGrid1.LayOut(2)
    End If

End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+----------------------------+
'+      取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 2            ' [受注]契約先

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+        品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 2            ' [受注]契約先
    RVI_Torcd = RTrim(iH1_Torcd)

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
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
    If TorNmGet(iH1_Torcd, 2) Then
       H1lb_Tornm = TRM_RName
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
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT START
'+-----------------------------+
'+  H0Op_Sort
'+-----------------------------+
Private Sub H0Op_Sort_Click(Index As Integer)
    H0Op_Sort(0).BackColor = mColor_Off
    H0Op_Sort(1).BackColor = mColor_Off
    H0Op_Sort(2).BackColor = mColor_Off
    
    H0Op_Sort(Index).BackColor = mColor_On
    Dsp_Flg = Index
End Sub
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT END
'+-----------------------------+
'+  H1Op_Dsp
'+-----------------------------+
Private Sub H1Op_Dsp_Click(Index As Integer)
    Select Case Index
        Case 0
            H1Op_Dsp(0).BackColor = mColor_On
            H1Op_Dsp(1).BackColor = mColor_Off
            H1Op_Dsp(2).BackColor = mColor_Off
        Case 1
            H1Op_Dsp(0).BackColor = mColor_Off
            H1Op_Dsp(1).BackColor = mColor_On
            H1Op_Dsp(2).BackColor = mColor_Off
        Case 2
            H1Op_Dsp(0).BackColor = mColor_Off
            H1Op_Dsp(1).BackColor = mColor_Off
            H1Op_Dsp(2).BackColor = mColor_On
    End Select

    If H1Op_Dsp(1).Value Then
        H1Op_Data(0).Enabled = True
        H1Op_Data(1).Enabled = True
        H1Op_Data(2).Enabled = True
        Data_Back.Visible = True
        Data_Back1.BackColor = &HC0C0FF
    End If
    If H1Op_Dsp(2).Value Then
        H1Op_Data(0).Value = True
        H1Op_Data(0).Enabled = False
        H1Op_Data(1).Enabled = False
        H1Op_Data(2).Enabled = False
        Data_Back.Visible = False
        Data_Back1.BackColor = &HC0C0C0
    End If
    If H1Op_Dsp(0).Value Then
        H1Op_Data(0).Enabled = True
        H1Op_Data(1).Enabled = True
        H1Op_Data(2).Enabled = True
        Data_Back.Visible = True
        Data_Back1.BackColor = &HC0C0FF
    End If
End Sub
'+-----------------------------+
'+  H1Op_Data
'+-----------------------------+
Private Sub H1Op_Data_Click(Index As Integer)
    Select Case Index
        Case 0
            H1Op_Data(0).BackColor = mColor_On
            H1Op_Data(1).BackColor = mColor_Off
            H1Op_Data(2).BackColor = mColor_Off
        Case 1
            H1Op_Data(0).BackColor = mColor_Off
            H1Op_Data(1).BackColor = mColor_On
            H1Op_Data(2).BackColor = mColor_Off
        Case 2
            H1Op_Data(0).BackColor = mColor_Off
            H1Op_Data(1).BackColor = mColor_Off
            H1Op_Data(2).BackColor = mColor_On
    End Select
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
     Case "DEL"
          SyoriNM = "削除"
          K_Sykbnm.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
''      * PF-Key 表示設定 *
'       If SyoriKB <> "LOC" Then
'          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
'             gObj.Name = "PB_DEL" Then
'             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
'          End If
'       End If
   Next gObj
   Timer2.Enabled = False
   If SyoriNM = "" Then
      K_Sykbnm.Visible = False
      K_Sykbnm.Caption = ""
      Exit Sub
   End If
   K_Sykbnm.Caption = SyoriNM
   K_Sykbnm.Visible = True
   Timer2.Enabled = True
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel
       Case "B1": B1Mode = True: B1Color = gPGotSel
    End Select
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT START
    H0_Area1.Enabled = H1Mode: H0_Area1.BackColor = H1Color
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT END
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color

    PB_Look.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_CHG.Visible = B1Mode
    PB_DSP.Visible = False
    PB_DEL.Visible = H1Mode

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
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = ""
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
        Call CisVsGrid1.Clear
        iB1_Nai1.Visible = False
        iB1_Nai2.Visible = False
        iB1_Nai3.Visible = False
    End If
End Sub
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        If SyoriKB <> "DEL" Then
           vsGrid1.HighLight = flexHighlightNever
           Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
        Else
           vsGrid1.SetFocus
           vsGrid1.HighLight = flexHighlightAlways
        End If
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT START
        With vsGrid1
            .Redraw = False
            
            For gInt = 1 To 11
                .Col = gInt
                .CellBackColor = mColorSel
            Next gInt
            .Redraw = True
        End With
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT END
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        If SyoriKB = "DEL" Then
           vsGrid1.HighLight = flexHighlightNever
        End If
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
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

'   内示年月
    If RTrim(iH1_Ym) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    内示年月必須入力！！       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ym.SetFocus
        Exit Function
    Else
        If Not iH1_Ym.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    年月を正しく入力して下さい      "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Ym.SetFocus
            Exit Function
        End If
    End If
'   取引先
    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 2) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    契約先　未登録！（取引先マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    Else
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先必須入力！！       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Torcd.SetFocus
        Exit Function
    End If

' 抽出中メッセージ表示
    
    PNL_Wait.Visible = True
    DoEvents
    
    Call SqlSelectSet
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        PNL_Wait.Visible = False
        DoEvents
        Exit Function
    End If

    PNL_Wait.Visible = False
    DoEvents
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False

    If SyoriKB <> "DEL" Then
        Call CisVsGrid1.EditEnable
        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
    End If

    Dim ID          As Long
    Dim wChk        As Byte
    
    wChk = 0
    With vsGrid1
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For

            If SyoriKB <> "DEL" Then
                If RTrim(.TextMatrix(ID, 18)) = "X" Then
                   wChk = 1
                End If
            Else
                If RTrim(.TextMatrix(ID, 19)) = "X" Then
                   wChk = 1
                End If
            End If
        Next ID
    End With
    
    If wChk = 0 Then
        With CisFun
            .MB_Lines = 4
            If SyoriKB <> "DEL" Then
               .MB_MSG(2) = "    データが変更されていません   "
            Else
               .MB_MSG(2) = "    削除データが選択されていません   "
            End If
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
            
            If SyoriKB = "DEL" Then vsGrid1.SetFocus
            Exit Function
        End With
    End If

    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    
    If GridSet1 Then
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'****************************
'*      SQL文ｾｯﾄ処理
'****************************
Private Sub SqlSelectSet(Optional InputKey As String = "")
    Dim wZenYM      As String
'
    '画面ワーク内容削除
    gSL_Select = "DELETE FROM 内示一覧画面ワーク"
    gSL_Select = gSL_Select & " WHERE 処理端末 = substring(host_name(),1,20)"
    With CisDB
        .SQL = gSL_Select
        If Not .DBExec Then Exit Sub
    End With
    
    '前年月取得
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add iH1_Ym, M, -1
        wZenYM = .Date_Result
    End With
'
    '------ 2010/01/26 INSERT START ------
    With CisFun
        .Where_CreateKB = NewDocument
        Call .Where_Create(AlphaMe, "NJ.内示年月", RTrim(iH1_Ym))
        Call .Where_Create(AlphaMe, "NJ.契約先", RTrim(iH1_Torcd))
        Call .Where_Create(AlphaMe, "NJ.品番", RTrim(iH1_Hinbn))
    End With
    '------ 2010/01/26 INSERT END   ------

    '基本SQL分セット
    gSL_Select = ""
    gSL_Select = gSL_Select & " INSERT INTO 内示一覧画面ワーク "
    '内示テーブル基本情報
    gSL_Select = gSL_Select & " SELECT substring(host_name(),1,20),NJ.品番,NJ.納入先,NJ.受入,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HN.表示品番,'') <> '' AND ISNULL(HT.品番,'') <> '' THEN HN.表示品番 ELSE NJ.品番 END 表示品番,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(TR.取引先CD,'') <> '' THEN RTRIM(TR.取引先CD) + ':' + RTRIM(TR.略称) ELSE NJ.納入先 END 表示納入先,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(NJ.背番号,'') <> '' THEN NJ.背番号 ELSE"
    gSL_Select = gSL_Select & "     CASE WHEN ISNULL(HT.背番号,'') <> '' THEN HT.背番号 ELSE ISNULL(ZN.背番号,'') END"
    gSL_Select = gSL_Select & " END 背番号,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(NJ.収容数,0) <> 0 THEN NJ.収容数 ELSE"
    gSL_Select = gSL_Select & "     CASE WHEN ISNULL(HT.変更日,'') = '' THEN"
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL(HT.収容数1,0) <> 0 THEN HT.収容数1"
    gSL_Select = gSL_Select & "             ELSE ISNULL(ZN.収容数,0) END"
    gSL_Select = gSL_Select & "     Else"
    gSL_Select = gSL_Select & "         CASE WHEN SUBSTRING(HT.変更日,1,6) > NJ.内示年月    THEN ISNULL(HT.収容数1,0) ELSE ISNULL(HT.収容数2,0) END"
    gSL_Select = gSL_Select & "     END END 収容数,"
    gSL_Select = gSL_Select & " ZN.内示数 前月当月内示,ZN.翌月内示数 前月翌月内示,"
    gSL_Select = gSL_Select & " NJ.内示数 , NJ.翌月内示数, NJ.翌々月内示数,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HT.品番,'') <> '' THEN '1' ELSE '' END 品番登録済FLG,"
    gSL_Select = gSL_Select & " '*' 内示登録FLG"
    gSL_Select = gSL_Select & " FROM 内示テーブル NJ"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & " ON   HN.品番 = NJ.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " ON   HT.品番 = NJ.品番"
    gSL_Select = gSL_Select & " AND HT.取引先 = NJ.納入先"
    gSL_Select = gSL_Select & " AND HT.受入 = NJ.受入"
    gSL_Select = gSL_Select & " AND HT.品目 = 0"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & " ON   TR.取引先CD = NJ.納入先"
    gSL_Select = gSL_Select & " AND TR.取引先区分 = 0"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 内示テーブル ZN"
    gSL_Select = gSL_Select & " ON   ZN.内示年月 = '" & RTrim(wZenYM) & "'"
    gSL_Select = gSL_Select & " AND ZN.品番 = NJ.品番"
    gSL_Select = gSL_Select & " AND ZN.納入先 = NJ.納入先"
    gSL_Select = gSL_Select & " AND ZN.受入 = NJ.受入"
    gSL_Select = gSL_Select & " AND ZN.内示区分 = NJ.内示区分 "
'------ 2010/01/26 UPDATE START ------
'    gSL_Select = gSL_Select & " WHERE NJ.内示年月 = '" & RTrim(iH1_Ym) & "'"
'    gSL_Select = gSL_Select & " AND   NJ.契約先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
'------ 2010/01/26 UPDATE END   ------
    '内示データ未登録のみ表示
    If H1Op_Dsp(1).Value Then
        gSL_Select = gSL_Select & " AND ISNULL(NJ.品番,'') = '' "
    End If
    '前月内示存在のみ表示
    If H1Op_Data(1).Value Then
        gSL_Select = gSL_Select & " AND ISNULL(ZN.品番,'') <> '' "
    End If
    '品番マスタ登録済のみ表示
    If H1Op_Data(2).Value Then
        gSL_Select = gSL_Select & " AND ISNULL(HT.品番,'') <> '' "
    End If
    
    If Not H1Op_Dsp(2).Value Then               '全てor内示未登録
        gSL_Select = gSL_Select & " UNION"
        
        '------ 2010/01/26 INSERT START ------
        With CisFun
            Call .Where_Create(AlphaMe, "HT.品番", RTrim(iH1_Hinbn))
        End With
        '------ 2010/01/26 INSERT END   ------
        
        '品番取引先マスタ基本情報
        gSL_Select = gSL_Select & " SELECT"
        gSL_Select = gSL_Select & " substring(host_name(),1,20),HT.品番,HT.取引先 納入先,HT.受入,"
        gSL_Select = gSL_Select & " CASE WHEN ISNULL(HN.表示品番,'') <> '' THEN HN.表示品番 ELSE HT.品番 END 表示品番,"
        gSL_Select = gSL_Select & " CASE WHEN ISNULL(TR.取引先CD,'') <> '' THEN RTRIM(TR.取引先CD) + ':' + RTRIM(TR.略称) ELSE HT.取引先 END 表示納入先,"
        gSL_Select = gSL_Select & " HT.背番号,"
        gSL_Select = gSL_Select & " CASE WHEN ISNULL(HT.変更日,'') = '' THEN"
        gSL_Select = gSL_Select & "     CASE WHEN ISNULL(HT.収容数1,0) <> 0 THEN HT.収容数1 ELSE 0 END"
        gSL_Select = gSL_Select & "     ELSE CASE WHEN SUBSTRING(HT.変更日,1,6) > '" & RTrim(iH1_Ym) & "'" & " THEN  ISNULL(HT.収容数1,0) ELSE ISNULL(HT.収容数2,0) END"
        gSL_Select = gSL_Select & " END 収容数,"
        gSL_Select = gSL_Select & " ZN.内示数 前月当月内示,ZN.翌月内示数 前月翌月内示,"
        gSL_Select = gSL_Select & " 0 内示数,0 翌月内示数,0 翌々月内示数,'1' 品番登録済FLG,"
        gSL_Select = gSL_Select & " '' 内示登録FLG"
        gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
        gSL_Select = gSL_Select & " ON   HN.品番 = HT.品番"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
        gSL_Select = gSL_Select & " ON   TR.取引先CD = HT.取引先"
        gSL_Select = gSL_Select & " AND TR.取引先区分 = 0"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 内示テーブル NJ"
        gSL_Select = gSL_Select & " ON   NJ.内示年月 = '" & RTrim(iH1_Ym) & "'"
        gSL_Select = gSL_Select & " AND NJ.品番 = HT.品番"
        gSL_Select = gSL_Select & " AND NJ.納入先 = HT.取引先"
        gSL_Select = gSL_Select & " AND NJ.受入 = HT.受入"
        gSL_Select = gSL_Select & " AND NJ.内示区分 = ''"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 内示テーブル ZN"
        gSL_Select = gSL_Select & " ON   ZN.内示年月 = '" & RTrim(wZenYM) & "'"
        gSL_Select = gSL_Select & " AND ZN.品番 = HT.品番"
        gSL_Select = gSL_Select & " AND ZN.納入先 = HT.取引先"
        gSL_Select = gSL_Select & " AND ZN.受入 = HT.受入"
        gSL_Select = gSL_Select & " AND ZN.内示区分 = ''"
        gSL_Select = gSL_Select & " WHERE HT.契約先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " AND   HT.品目 = 0 "
        gSL_Select = gSL_Select & " AND   ISNULL(NJ.品番,'') = '' "
        '------ 2010/01/26 INSERT START ------
        gSL_Select = gSL_Select & CisFun.Where_ThisPhrase
        '------ 2010/01/26 INSERT END   ------
    
        '前月内示存在のみ表示
        If H1Op_Data(1).Value Then
            gSL_Select = gSL_Select & " AND ZN.品番 <> '' "
        End If
    
    End If

    If (H1Op_Dsp(0).Value = True Or H1Op_Dsp(1).Value = True) And (H1Op_Data(1).Value = True Or H1Op_Data(0).Value = True) Then                   '前月内示＆品番未登録対象
        gSL_Select = gSL_Select & " UNION"
        
        '------ 2010/01/26 INSERT START ------
        With CisFun
            Call .Where_Create(AlphaMe, "NJ.品番", RTrim(iH1_Hinbn))
        End With
        '------ 2010/01/26 INSERT END   ------
        
        '前月内示テーブル基本情報
        gSL_Select = gSL_Select & " SELECT substring(host_name(),1,20),NJ.品番,NJ.納入先,NJ.受入,"
        gSL_Select = gSL_Select & " NJ.品番 表示品番,"
        gSL_Select = gSL_Select & " CASE WHEN ISNULL(TR.取引先CD,'') <> '' THEN RTRIM(TR.取引先CD) + ':' + RTRIM(TR.略称) ELSE NJ.納入先 END 表示納入先,"
        gSL_Select = gSL_Select & " NJ.背番号 背番号,"
        gSL_Select = gSL_Select & " NJ.収容数 収容数,"
        gSL_Select = gSL_Select & " NJ.内示数 前月当月内示,NJ.翌月内示数 前月翌月内示,"
        gSL_Select = gSL_Select & " 0 内示数, 0 翌月内示数, 0 翌々月内示数,"
        gSL_Select = gSL_Select & " '' 品番登録済FLG,"
        gSL_Select = gSL_Select & " '' 内示登録FLG"
        gSL_Select = gSL_Select & " FROM 内示テーブル NJ"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
        gSL_Select = gSL_Select & " ON   HT.品番 = NJ.品番"
        gSL_Select = gSL_Select & " AND HT.取引先 = NJ.納入先"
        gSL_Select = gSL_Select & " AND HT.受入 = NJ.受入"
        gSL_Select = gSL_Select & " AND HT.品目 = 0"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
        gSL_Select = gSL_Select & " ON   TR.取引先CD = NJ.納入先"
        gSL_Select = gSL_Select & " AND TR.取引先区分 = 0"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 内示テーブル ZN"
        gSL_Select = gSL_Select & " ON   ZN.内示年月 = '" & RTrim(iH1_Ym) & "'"
        gSL_Select = gSL_Select & " AND ZN.品番 = NJ.品番"
        gSL_Select = gSL_Select & " AND ZN.納入先 = NJ.納入先"
        gSL_Select = gSL_Select & " AND ZN.受入 = NJ.受入"
        gSL_Select = gSL_Select & " AND ZN.内示区分 = NJ.内示区分"
        gSL_Select = gSL_Select & " WHERE NJ.内示年月 = '" & RTrim(wZenYM) & "'"
        gSL_Select = gSL_Select & " AND   NJ.契約先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " AND ISNULL(HT.品番,'') = '' "
        gSL_Select = gSL_Select & " AND ISNULL(ZN.品番,'') = '' "
        '------ 2010/01/26 INSERT START ------
        gSL_Select = gSL_Select & CisFun.Where_ThisPhrase
        '------ 2010/01/26 INSERT END   ------
    
    End If
    
    '内示一覧画面ワーク出力
    With CisDB
        .SQL = gSL_Select
        If Not .DBExec Then Exit Sub
    End With
    
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT * FROM 内示一覧画面ワーク "
    gSL_Select = gSL_Select & " WHERE 処理端末 = substring(host_name(),1,20)"
    Select Case Dsp_Flg
           Case 0
                gSL_Select = gSL_Select & " ORDER BY 品番,納入先,受入,背番号 "
           Case 1
                gSL_Select = gSL_Select & " ORDER BY 背番号,品番,納入先,受入 "
           Case 2
                gSL_Select = gSL_Select & " ORDER BY 納入先,受入,品番,背番号 "
    End Select
'
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False

    With CisFun
        .MB_Lines = 5
        If SyoriKB <> "DEL" Then
           .MB_MSG(2) = "    更新処理を行います。            "
        Else
           .MB_MSG(2) = "    削除処理を行います。            "
        End If
        .MB_MSG(4) = "     よろしいですか？            "
        If SyoriKB <> "DEL" Then
           .MB_Title = "更新処理"
        Else
           .MB_Title = "削除処理"
        End If
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

'    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 更新処理 )
    Dim ID          As Long
    If SyoriKB <> "DEL" Then
        With vsGrid1
            For ID = 1 To .Rows - 1
                If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
'                If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(7))) <> 0 Or CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(8))) <> 0 Or _
'                   CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(9))) <> 0 Then
                     If RTrim(.TextMatrix(ID, 18)) = "X" Then Call DBUpdate_NJT(ID)
'                End If
            Next ID
        End With
    End If
' ( 削除処理 )
    If SyoriKB = "DEL" Then
        With vsGrid1
            For ID = 1 To .Rows - 1
                If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
                If RTrim(.TextMatrix(ID, 19)) = "X" Then Call DBDelete_NJT(ID)
            Next ID
        End With
    End If

'    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      内示テーブル追加・更新
'+------------------------------------+
Public Function DBUpdate_NJT(SetRow As Long, Optional DBNo As Integer) As Boolean
    DBUpdate_NJT = False
    
    Up_Flg = 0
    '内示テーブルデータ存在確認
    gSL_Select = ""
    gSL_Select = "SELECT * FROM 内示テーブル "
    gSL_Select = gSL_Select & " WHERE 内示年月 = '" & RTrim(iH1_Ym) & "'"
    gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(12))) & "'"
    gSL_Select = gSL_Select & " AND 納入先 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(13))) & "'"
    gSL_Select = gSL_Select & " AND 受入 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(3))) & "'"
    gSL_Select = gSL_Select & " AND 内示区分 = ''"
    
    If NJTRead(gSL_Select, 1) Then
       Up_Flg = 1                       '更新
    Else
       Up_Flg = 0                       '追加
    End If
    
    If Up_Flg = 1 Then
        With NJT
            Call DayReCalc(SetRow)          '日量数算出
            
            gSL_Select = "UPDATE 内示テーブル SET "
            gSL_Select = gSL_Select & " 内示数 = " & CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(8))) & ","
            gSL_Select = gSL_Select & " 翌月内示数 = " & CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(9))) & ","
            gSL_Select = gSL_Select & " 翌々月内示数 = " & CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(10))) & ","
            For gInt = 1 To 31
                gSL_Select = gSL_Select & " 内示数量" & Format(gInt, "00") & " = " & .内示数量(gInt) & ","
            Next gInt
            gSL_Select = gSL_Select & " 登録区分 = 1,"
            gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
            gSL_Select = gSL_Select & " 更新日 = getdate(),"
            gSL_Select = gSL_Select & " 更新端末 = substring(host_name(),1,20)"
            gSL_Select = gSL_Select & " WHERE 内示年月 = '" & RTrim(iH1_Ym) & "'"
            gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(12))) & "'"
            gSL_Select = gSL_Select & " AND 納入先 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(13))) & "'"
            gSL_Select = gSL_Select & " AND 受入 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(3))) & "'"
            gSL_Select = gSL_Select & " AND 内示区分 = ''"
        
            With CisDB
                .ConnectNo = DBNo
                .SQL = gSL_Select
                If Not .DBExec Then Exit Function
            End With
    
        End With
    Else
        Call ItemsClearNJT
        Call DayReCalc(SetRow)          '日量数算出
        With NJT
            .内示年月 = RTrim(iH1_Ym)
            .品番 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(12)))
            .納入先 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(13)))
            .受入 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(3)))
            .契約先 = RTrim(iH1_Torcd)
            .内示数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(8)))
            .翌月内示数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(9)))
            .翌々月内示数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(10)))
            .収容数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(5)))
            .背番号 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(4)))
            .登録区分 = 0
            .作成者 = gTanto
        End With
        Call NJTInsert
    End If
    
' オンラインデータ以外の時、内示前回テーブルへ出力
    With CisDB
         .SQL = "Delete 内示前回テーブル"
         .SQL = .SQL & " Where 内示年月 = '" & RTrim(iH1_Ym) & "'"
         .SQL = .SQL & " And   品番 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(12))) & "'"
         .SQL = .SQL & " And   納入先 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(13))) & "'"
         .SQL = .SQL & " And   受入 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(3))) & "'"
         .SQL = .SQL & " And   内示区分 = ''"
         .DBExec
         
         .SQL = "Insert Into 内示前回テーブル "
         .SQL = .SQL & " Select * From 内示テーブル "
         .SQL = .SQL & " Where 内示年月 = '" & RTrim(iH1_Ym) & "'"
         .SQL = .SQL & " And   品番 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(12))) & "'"
         .SQL = .SQL & " And   納入先 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(13))) & "'"
         .SQL = .SQL & " And   受入 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(3))) & "'"
         .SQL = .SQL & "    And   内示区分 = ''"
         .DBExec
    End With
    
    DBUpdate_NJT = True
End Function
'+-----------------------------------------+
'+      日量数計算 内示テーブルセット
'+-----------------------------------------+
Private Sub DayReCalc(SetRow As Long)
    
    With CisDB
        .SQL = "日量数算出"
        .StoadoCount = 35
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用

        For gLong = 5 To 35
            .ParaNo = gLong: .ParaIO = OutPut用
        Next gLong

        .ParaNo = 1:    .ParaValue = iH1_Ym
        .ParaNo = 2:    .ParaValue = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(8)))
        .ParaNo = 3:    .ParaValue = RTrim$(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(13)))
        .ParaNo = 4:    .ParaValue = RTrim$(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(3)))

        .DBStored

        For gInt = 5 To 35
            .ParaNo = gInt
            gLong = gInt - 5
            NJT.内示数量(gLong + 1) = CisFun.Val2(.ParaValue)
        Next gInt
    
    End With

End Sub
'+------------------------------------+
'+      内示テーブル削除
'+------------------------------------+
Public Function DBDelete_NJT(SetRow As Long, Optional DBNo As Integer) As Boolean
    DBDelete_NJT = False
            
    gSL_Select = "DELETE FROM 内示テーブル "
    gSL_Select = gSL_Select & " WHERE 内示年月 = '" & RTrim(iH1_Ym) & "'"
    gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(12))) & "'"
    gSL_Select = gSL_Select & " AND 納入先 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(13))) & "'"
    gSL_Select = gSL_Select & " AND 受入 = '" & RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(3))) & "'"
    gSL_Select = gSL_Select & " AND 内示区分 = ''"

    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    
    DBDelete_NJT = True
End Function

'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 4, 0, "№")
'        Call .FixedSet(中中, 左中, あり, 27, 1, "品　番")          '2010/07/07
        Call .FixedSet(中中, 左中, あり, 25, 1, "品　番")           '2010/07/07
        Call .FixedSet(中中, 左中, あり, 20, 2, "納　入　先")
        Call .FixedSet(中中, 中中, あり, 2, 3, "受")
'        Call .FixedSet(中中, 左中, あり, 6, 4, "背番号")           '2010/07/07
        Call .FixedSet(中中, 左中, あり, 8, 4, "背番号")            '2010/07/07
        Call .FixedSet(中中, 右中, あり, 6, 5, "収容数")
        Call .FixedSet(中中, 右中, あり, 16, 6, "［前月］当月内示")
        Call .FixedSet(中中, 右中, あり, 16, 7, "［前月］翌月内示")
        Call .FixedSet(中中, 右中, あり, 10, 8, "当月内示　")
        Call .FixedSet(中中, 右中, あり, 10, 9, "翌月内示　")
        Call .FixedSet(中中, 右中, あり, 10, 10, "翌々月内示")
        Call .FixedSet(中中, 左中, あり, 1, 11, " ")
        Call .FixedSet(中中, 右中, あり, 0, 12, "X品番")
        Call .FixedSet(中中, 左中, あり, 0, 13, "X納入先CD")
        Call .FixedSet(中中, 右中, あり, 0, 14, "X当月内示")
        Call .FixedSet(中中, 右中, あり, 0, 15, "X翌月内示")
        Call .FixedSet(中中, 左中, あり, 0, 16, "X翌々月内示")
        Call .FixedSet(中中, 左中, なし, 0, 17, "X品番ﾏｽﾀ登録済ﾌﾗｸﾞ")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 18, "Select")      '更新用Select
        Call .FixedSet(中中, 左中, なし, 0, 19, "Select")      '削除用Select
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,8,9,10,11")       'パターン①[品番/納入先/受入/背番号]
        Call .PatanSet(1, "0,1,2,3,4,5,7,8,9,10,11")       'パターン①-1[前月翌月内示]
        Call .PatanSet(2, "0,2,3,1,4,5,6,8,9,10,11")       'パターン②[納入先/受入/品番/背番号]
        Call .PatanSet(3, "0,2,3,1,4,5,7,8,9,10,11")       'パターン②-1[前月翌月内示]
    
'   編集内容ｾｯﾄ
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO,非表示ｶﾗﾑ,編集ﾌｫｰﾏｯﾄ,日付編集)
        Call .EditSet(iB1_Nai1, 7, , "###,###")
        Call .EditSet(iB1_Nai2, 8, , "###,###")
        Call .EditSet(iB1_Nai3, 9, , "###,###")
'        Call .EditSet(iB1_BinSu, 10, 17)
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット
'***********************************************'
Private Function GridSet1() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet1 = False

    With CisDB
        .SQL = gSL_Select
        
        If .DBRead Then
    
    '       抽出処理 初期ｾｯﾄ
            B1lb_Kensu = Format(.RecordCount, "#,##0")
            Ttl_Kensu = .RecordCount       ' 抽出件数(合計)退避
    '       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
            With vsGrid1
               .Redraw = False
                If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
                    .Rows = CisVsGrid1.DispMax
                Else
                    .Rows = Ttl_Kensu + 1
                End If
            End With
            ID = 0
            Do Until .Recordset.EOF
                ID = ID + 1
    
                Call DBInfoToGrid1(ID)
    
                .Recordset.MoveNext
            Loop
            With vsGrid1
                 .Row = 1: .Col = 1: .ColSel = .Cols - 1
                 .Redraw = True
            End With
            NowRow = vsGrid1.Row
    
            GridSet1 = True
        End If
        
        .Recordset.Close
        
    End With

End Function
'+----------------------------------------------+
'+  ＤＢ内容→グリッドに転送
'+----------------------------------------------+
Private Sub DBInfoToGrid1(SetRow As Long)
    With vsGrid1
        'No
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(0)) = SetRow
        '表示品番
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = CisDB.Recordset("表示品番")
        '納入先
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = CisDB.Recordset("表示納入先")
        '受入
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = CisDB.Recordset("受入")
        '背番号
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = CisDB.Recordset("背番号")
        '収容数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = CisDB.Recordset("収容数")
        '前月[当月数量]
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = Format(CisDB.Recordset("前月当月内示"), "##,###")
        '前月[翌月数量]
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = Format(CisDB.Recordset("前月翌月内示"), "##,###")
        '当月内示
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = Format(CisDB.Recordset("内示数"), "##,###")
        '翌月内示
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = Format(CisDB.Recordset("翌月内示数"), "##,###")
        '翌々月内示
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(10)) = Format(CisDB.Recordset("翌々月内示数"), "##,###")
        '内示登録FLG
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(11)) = CisDB.Recordset("内示登録FLG")
        '退避
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(12)) = CisDB.Recordset("品番")
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(13)) = CisDB.Recordset("納入先")
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(14)) = CisDB.Recordset("内示数")
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(15)) = CisDB.Recordset("翌月内示数")
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(16)) = CisDB.Recordset("翌々月内示数")
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(17)) = CisDB.Recordset("品番登録済FLG")
    End With
End Sub
'+-------------------------------+
'+      VSGrid1_Click
'+-------------------------------+
Private Sub VSGrid1_Click()
    If ProcHB = "B1" Then

        If SyoriKB <> "DEL" Then
            If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(0))) = "" Then
                vsGrid1.Row = NowRow
            End If
    
            Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
            Call CisVsGrid1.EditEnable
            Call CisVsGrid1.EditVisible
            If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(0))) <> "" Then
                NowRow = vsGrid1.Row
            End If
        Else
            
            If vsGrid1.Row > vsGrid1.RowSel Then
               gInt = vsGrid1.Row
               vsGrid1.Row = vsGrid1.RowSel
               vsGrid1.RowSel = gInt
            End If
            For wInt = vsGrid1.Row To vsGrid1.RowSel
                If RTrim(vsGrid1.TextMatrix(wInt, 0)) <> "" Then
                    vsGrid1.Row = wInt
                    If Trim(vsGrid1.TextMatrix(vsGrid1.Row, 19)) = "" Then
                        vsGrid1.FillStyle = 1
                        vsGrid1.Col = 1: vsGrid1.ColSel = 11
                        vsGrid1.CellBackColor = mColorDel
                        vsGrid1.FillStyle = 0
                        vsGrid1.TextMatrix(vsGrid1.Row, 19) = "X"
                        vsGrid1.Col = 0
                        vsGrid1.CellBackColor = mColorDel
                    Else
                        vsGrid1.FillStyle = 1
                        vsGrid1.Col = 1: vsGrid1.ColSel = 11
                        vsGrid1.CellBackColor = mColorOrg
                        vsGrid1.FillStyle = 0
                        vsGrid1.TextMatrix(vsGrid1.Row, 19) = ""
                        vsGrid1.Col = 0
                        vsGrid1.CellBackColor = &HE0E0E0
                    End If
                End If
            Next wInt
        End If
        
        
    End If
End Sub
'+---------------------------------------------------------+
'+          グリッド編集エリア
'+---------------------------------------------------------+
'+-------------------------------+
'+      iB1_Nai1
'+-------------------------------+
Private Sub iB1_Nai1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyTab Then
        If Shift = 0 Then
           iB1_Nai2.SetFocus
        Else
            Call BD_InKeyCheck(ActiveControl, vbKeyUp, Shift)
            iB1_Nai3.SetFocus
        End If
    End If
    
    Call BD_InKeyCheck(ActiveControl, KeyCode, Shift)

End Sub
'+-------------------------------+
'+      iB1_Nai2
'+-------------------------------+
Private Sub iB1_Nai2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyTab Then
        If Shift = 0 Then
            iB1_Nai3.SetFocus
        Else
            iB1_Nai1.SetFocus
        End If
    End If
    
    Call BD_InKeyCheck(ActiveControl, KeyCode, Shift)

End Sub
'+-------------------------------+
'+      iB1_Nai3
'+-------------------------------+
Private Sub iB1_Nai3_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyTab Then
       Call BD_InKeyCheck(ActiveControl, KeyCode, Shift)
    Else
        If Shift = 0 Then
            Call BD_InKeyCheck(ActiveControl, vbKeyDown, Shift)
            iB1_Nai1.SetFocus
        Else
            Call BD_InKeyCheck(ActiveControl, KeyCode, Shift)
            iB1_Nai2.SetFocus
        End If
    End If

End Sub
'+--------------------------------------+
'+      ボディ入力項目キーチェック
'+--------------------------------------+
Private Function BD_InKeyCheck(TObj As Object, KeyCode As Integer, Shift As Integer)

'< PageUp・PageDown >
    If KeyCode = vbKeyPageDown Or _
       KeyCode = vbKeyPageUp Then
        
        If vsGrid1.Rows <= CisVsGrid1.DispMax Then Exit Function
        
        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        
        If KeyCode = vbKeyPageDown Then
            
            '< PageDown >
            If vsGrid1.TopRow + 25 <= vsGrid1.Rows - 1 Then
                vsGrid1.TopRow = vsGrid1.TopRow + 25
                
                If vsGrid1.TopRow + 24 <= vsGrid1.Rows - 1 Then
                    vsGrid1.Row = vsGrid1.TopRow + 24
                Else
                    vsGrid1.Row = vsGrid1.Rows - 1
                End If
            Else
                vsGrid1.TopRow = vsGrid1.Rows - 1 - 24
            End If
        Else
            
            '< PageUp >
            If vsGrid1.TopRow - 25 > 0 Then
                vsGrid1.TopRow = vsGrid1.TopRow - 25
            Else
                vsGrid1.TopRow = 1
            End If
            vsGrid1.Row = vsGrid1.TopRow
        End If
        
        vsGrid1.SetFocus
        Call CisVsGrid1.EditEnable
        Call CisVsGrid1.EditVisible
        NowRow = vsGrid1.Row
        TObj.SetFocus
        Exit Function
    End If
    
'< ↑・↓ >
    If KeyCode = vbKeyUp Or _
       KeyCode = vbKeyDown Then
        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        vsGrid1.SetFocus
        Call CisVsGrid1.KeyEdit(TObj, KeyCode, Shift)
        NowRow = vsGrid1.Row
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT START
        With vsGrid1
            .Redraw = False
            
            For gInt = 1 To 11
                .Col = gInt
                .CellBackColor = mColorSel
            Next gInt
            .Redraw = True
        End With
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT END
    End If

End Function
'+-------------------------------+
'+      背景色の再設定
'+-------------------------------+
Private Sub ChgColorGrid1(SetRow As Long)

    With vsGrid1
                  
        .Redraw = False
        
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT START
        For gInt = 1 To 11
            .Col = gInt
            .CellBackColor = mColorSel
        Next gInt
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/03/12 INSERT END
        
        NewRow = .Row
        .Row = SetRow
        If CisFun.Val2(iB1_Nai1) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(14))) Or _
           CisFun.Val2(iB1_Nai2) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(15))) Or _
           CisFun.Val2(iB1_Nai3) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(16))) Then
            For gInt = 1 To 11
                .Col = gInt
                .CellBackColor = mColorChg
            Next gInt
            .TextMatrix(SetRow, 18) = "X"
        Else
            For gInt = 1 To 11
                .Col = gInt
                .CellBackColor = mColorOrg
            Next gInt
            .TextMatrix(SetRow, 18) = ""
        End If
        .Row = NewRow
        .Redraw = True
        
    End With
    
End Sub


