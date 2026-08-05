VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdodc.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CisYM_6.0.ocx"
Begin VB.Form CZA0410 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "在庫状況照会"
   ClientHeight    =   11010
   ClientLeft      =   3540
   ClientTop       =   1155
   ClientWidth     =   14895
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
   ScaleHeight     =   11010
   ScaleWidth      =   14895
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   5205
      Top             =   5985
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
   Begin VB.PictureBox B1_Area1 
      Height          =   7920
      Left            =   255
      ScaleHeight     =   7860
      ScaleWidth      =   14340
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2685
      Width           =   14400
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CZA0410.frx":0000
         Height          =   7740
         Left            =   45
         OleObjectBlob   =   "CZA0410.frx":0015
         TabIndex        =   19
         Top             =   45
         Width           =   14235
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1290
      Left            =   270
      ScaleHeight     =   1230
      ScaleWidth      =   9660
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1260
      Width           =   9720
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   585
         Left            =   1665
         Top             =   615
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   1032
         ForeColor       =   16711680
         Caption         =   "品　番"
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
         cAlingnment     =   1
         cPositionX      =   75
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   870
            TabIndex        =   2
            Top             =   120
            Width           =   3210
            _ExtentX        =   5662
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
            Text            =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
            MaxLength       =   25
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   570
         Left            =   1665
         Top             =   45
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "契約先"
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
         cAlingnment     =   1
         cPositionX      =   75
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   870
            TabIndex        =   1
            Top             =   120
            Width           =   990
            _ExtentX        =   1746
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   360
            Left            =   1860
            Top             =   120
            Width           =   2460
            _ExtentX        =   4339
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
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1155
         Index           =   0
         Left            =   60
         Top             =   45
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   2037
         BackColor       =   16761087
         ForeColor       =   16711680
         Caption         =   "対象年月"
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
         cAlingnment     =   6
         cPositionY      =   45
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   330
            Left            =   165
            TabIndex        =   0
            Top             =   525
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cChkResult      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   570
         Left            =   6105
         Top             =   45
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "加工度"
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
         cAlingnment     =   1
         cPositionX      =   75
         Begin VB.ComboBox H1Comb_KakoDo 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "CZA0410.frx":148B5
            Left            =   1200
            List            =   "CZA0410.frx":148B7
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   3
            Top             =   120
            Width           =   1695
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   585
         Left            =   6105
         Top             =   615
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1032
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   75
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
      Caption         =   "　　　【 在庫状況問合せ 】"
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
         Left            =   4665
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   1995
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
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
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   780
         End
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
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
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
   Begin VB.PictureBox Back_Dsp 
      BackColor       =   &H00404040&
      Height          =   720
      Left            =   270
      ScaleHeight     =   660
      ScaleWidth      =   5295
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   450
      Width           =   5355
      Begin Cis3D_v60.CIS3D UC_3D5 
         Height          =   645
         Left            =   15
         Top             =   0
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   1138
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
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D UC_3D9 
            Height          =   570
            Index           =   0
            Left            =   60
            Top             =   30
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   1005
            BackColor       =   16761024
            Caption         =   "表示切替"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D10 
            Height          =   570
            Left            =   735
            Top             =   30
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   1005
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "⑤"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   2730
               TabIndex        =   8
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "④"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   2070
               TabIndex        =   7
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "③"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   1395
               TabIndex        =   6
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "②"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   5
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H008080FF&
               Caption         =   "①"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   4
               Top             =   150
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   570
            Left            =   4260
            Top             =   30
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   1005
            ForeColor       =   16711935
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
            Begin VB.OptionButton Op_Sort 
               BackColor       =   &H00C0FFC0&
               Caption         =   "昇順"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   0
               Left            =   30
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   60
               Value           =   -1  'True
               Width           =   825
            End
            Begin VB.OptionButton Op_Sort 
               BackColor       =   &H00C0C0C0&
               Caption         =   "降順"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   1
               Left            =   30
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   315
               Width           =   825
            End
         End
      End
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Left            =   240
      TabIndex        =   13
      Text            =   "D"
      Top             =   210
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   510
      Top             =   210
   End
   Begin VB.Timer Timer2 
      Left            =   1230
      Top             =   210
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   705
      Left            =   8835
      ScaleHeight     =   645
      ScaleWidth      =   5970
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   450
      Width           =   6030
      Begin CisBtn_60.CisBtn PB_Set 
         Height          =   585
         Left            =   2580
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
         BtnType         =   15
         BtnCaption      =   "設定"
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
         PFCaption       =   "( F11 )"
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
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   5100
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
         Left            =   4260
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
         Left            =   3420
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
      Begin CisBtn_60.CisBtn PB_Csv 
         Height          =   585
         Left            =   1740
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
         BtnType         =   14
         BtnCaption      =   "出力"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   192
         BtnShadowColor  =   16777215
         PFCaption       =   "( F10 )"
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
         PFBackColor     =   192
      End
      Begin CisBtn_60.CisBtn PB_Prt 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   13
         BtnCaption      =   "印刷"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   65535
         BtnShadowColor  =   0
         PFCaption       =   "( F8 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   65535
      End
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   45
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   0
      Top             =   420
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   405
      Left            =   0
      Top             =   765
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
Attribute VB_Name = "CZA0410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   在庫状況問合せ
'**       フォームID    :   CZA0410 (Based CSJ0410)
'**       処理概要      :
'**
'**       作  成  日    :  2019/07/16  By CIS
'**       変  更  日    :  2020/02/10  By CIS  ・・・・
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean          ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2       ' 入力域切替
    Dim MaxCnt                  As Long
    Dim NowCnt                  As Long
    Dim mSL_SelectW             As String           ' SQL SELECT(Where)条件文

    Dim CisDBGrid               As New CisDBGrid    ' TDBGridPro60ｸﾗｽ定義
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long

    Dim CisComb1                As CisCombo         ' 2019/07/16 ADD
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()

    If Rv_Call3 = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        If Not IniGet Then End
    Else
        FormAct = False
    End If
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)

    mGotColor = Op_Sort(0).BackColor
    mLostColor = Op_Sort(1).BackColor
    
'+---------------------+
'+ TrueGrid60 初期設定  +
'+---------------------+
    With CisDBGrid
        Set .DBGridObj = DBGrid1
        .FrmCaption = Me.Caption
        .MaxSort = 7                        ' ﾚｲｱｳﾄｿｰﾄ最大数
        .LayOut = True                      ' ﾚｲｱｳﾄ変更処理 True:有り,False:無し
        .LayOutMax = 5                      ' ﾚｲｱｳﾄ情報最大数
        .LayOutColMax = 30                  ' ﾚｲｱｳﾄ表示可能項目最大数
        .LayOutPath = gIniExe               ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ場所
        .LayOutName = CisFun.WSID & "ZA0410.Lay"       ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ名    ' 2019/07/08 UPD："SJ0410.Lay"
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        .CsvName = Trim(Me.Caption) & ".csv"
        .CsvRegKey = gRegKey
        .CsvValue = "CZA0410"               ' 2019/07/17 UPDｌ"CSJ0410"
        .GridInitGet = False                 ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ
        .GridProcGet = False
        .Init
    End With
    Call GridInitSet
    
    '+---------------------+
    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
    '+---------------------+
    Set CisComb1 = New CisCombo
    With CisComb1
        Set .Connect = CisDB
        Set .CombObj = H1Comb_KakoDo
        .Control = "名称マスタ"
        .ComboName = "品番区分"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
        .ComboSet
    End With

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
  ''PNL_ExeMsg.Move 4600, 400               ' 2019/07/16 DEL
    Call Str_Monitor("Start")               ' 2019/07/08 ADD

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
' Head部初期化
    B1Op_Patn(0).Value = True
    Op_Sort(0).Value = True
    Op_Sort(0).BackColor = mGotColor
    
    Op_Overlap(0).Value = True
    Op_Overlap(0).BackColor = mGotColor
    
    Call HeadBodyClear("H")
    
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    Adodc2.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    Adodc2.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
    
    iH1_YM.Text = Format(Now(), "yyyyMM")   ' 2019/07/16 ADD
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
           Case vbKeyF10:       If PB_Csv.Visible Then Call PB_Csv_Click    '【出力】
           Case vbKeyF11:       If PB_Set.Visible Then Call PB_Set_Click    '【設定】
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
    If Rv_Call3 = "" Then
    '+---------------------+
    '+   ﾃﾞｰﾀﾍﾞｰｽ切断
    '+---------------------+
        Call CisDB.DBDISConnect
    '+---------------------+
    '+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    '+---------------------+
        Call CisFrm.UnLoadDisp
        End
    Else
      ''Unload CSJ0410                      ' 2019/07/16 DEL
        Unload CZA0410                      ' 2019/07/16 ADD
        If Not Rv_CallObj Is Nothing Then
           Rv_CallObj.Show
        End If
    End If
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
      ''iH1_SDate.SetFocus      ' 2019/07/16 DEL
        iH1_YM.SetFocus         ' 2019/07/16 ADD
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call DBGridClr
      ''iH1_SDate.SetFocus      ' 2019/07/16 DEL
        iH1_YM.SetFocus         ' 2019/07/16 ADD
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B2" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    CSV出力           　     +
'+----------------------------+
Private Sub PB_Csv_Click()
    Call GridInitSetCsv

    With CisDBGrid
        Set .Recordset = Adodc1.Recordset
        .PutCsv
    End With

    Call GridInitSet
End Sub
'+----------------------------+
'+    印刷             　     +
'+----------------------------+
Private Sub PB_Prt_Click()
    With CisDBGrid
        Set .Recordset = Adodc1.Recordset
        .PutPrt
    End With
End Sub
'+----------------------------+
'+    設定              　     +
'+----------------------------+
Private Sub PB_Set_Click()
    With CisDBGrid
        .LayOutPatan
        If .LayOutSetFlg Then
           B1Op_Patn(.LayOutSetNo).Value = True
           If ProcHB = "B1" Then
              Call DBInput("R")
           End If
        End If
    End With
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub

'+-----------------------------+
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
        
    RVI_Torcd = Trim(iH1_Torcd)
    RVI_Hinmoku = 1
    RVI_Tehai = 2
    RVI_SK = 9      '
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RV_Hinbn
    End If
End Sub

'+-----------------------------+
'+       取引先(社内)検索      +
'+-----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_TorKb2 = 1

    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RV_Torcd
       Call iH1_Torcd_LostFocus
    End If
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub

'===================================================' 2019/07/17 ADD START
'+-----------------------------+
'+  iH1_YM
'+-----------------------------+
Private Sub iH1_YM_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = False
End Sub
Private Sub iH1_YM_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  H1Comb_KakoDo
'+-----------------------------+
Private Sub H1Comb_KakoDo_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = False
End Sub
Private Sub H1Comb_KakoDo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'===================================================' 2019/07/17 ADD END

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
    
    H1lb_Torcd = ""
    If Trim(iH1_Torcd) <> "" Then
        gSL_Select = "SELECT * FROM 取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先CD = '" & Trim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 社内区分 = 1"
        If TRMRead(gSL_Select) Then
            H1lb_Torcd = TRM.正式名称
        Else
            H1lb_Torcd = ""
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel: PB_ENT.BtnCaption = "入力"
        Case "B1":      B1Mode = True:  B1Color = gPGotSel: PB_ENT.BtnCaption = "詳細"
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_Look.Visible = H1Mode
    PB_END.Visible = H1Mode
'==========================================' 2019/07/08 DEL START
'    If Not B1Mode Then
'       If PNL_ExeMsg.Visible Then
'          PNL_ExeMsg.Visible = False
'       End If
'    End If
'==========================================' 2019/07/08 DEL END
    PB_ENT.Visible = H1Mode             ' 2009.12.11 upd : PB_ENT.Visible = H1Mode Or B1Mode
    PB_Set.Visible = Not B1Mode
    PB_Prt.Visible = B1Mode
    PB_Csv.Visible = B1Mode
    
    If Rv_Call3 = "K" Then
       PB_END.BtnCaption = "戻る"
       PB_END.Visible = True
    End If
    
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
            If .Name Like "Op" & wbc_Enm & "*" Then
               .Value = False
               .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "*Comb*" Then .ListIndex = 0
        End With
    Next gObj
    
    If HeadBody = "H" Then
        iH1_YM = Format(Date, "YYYYMM")
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
       If Not Head1Chk Then
          GoTo ReturnPress_Ed
       End If
       ProcHB = "B1"
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

    If Trim(iH1_YM) = "" Or Not iH1_YM.cChkResult() Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    対象年月を正しく入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
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
'+----------------------------+
'+    表示切替ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　     +
'+----------------------------+
Private Sub B1Op_Patn_Click(INDEX As Integer)
    Call CisFrm.MousePT(11)

    For Each gObj In B1Op_Patn
        If INDEX = gObj.INDEX Then
           gObj.BackColor = &H8080FF
        Else
           gObj.BackColor = &HC0C0C0
        End If
    Next gObj

    CisDBGrid.LayOutNow = INDEX + 1
    CisDBGrid.LayOutDisp

    If ProcHB = "B1" Then
        Call DBInput("R")
    End If
    Call CisFrm.MousePT(1)

End Sub
'+---------------------------+
'+    重複表示　ﾁｪｯｸﾎﾞｯｸｽ     +
'+---------------------------+
Private Sub Op_Overlap_GotFocus(INDEX As Integer)
    For Each gObj In Op_Overlap
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub Op_Overlap_Click(INDEX As Integer)

    If Not FormAct Then Exit Sub
    
    Call CisFrm.MousePT(11)

    Call CisDBGrid.MargeDisp(Op_Overlap(1).Value)

    If ProcHB = "H1" Then
      ''iH1_SDate.SetFocus      ' 2019/07/16 DEL
        iH1_YM.SetFocus         ' 2019/07/16 ADD
    Else
        DBGrid1.SetFocus
    End If

    Call CisFrm.MousePT(1)
End Sub
'+---------------------------+
'+    分類表示　ﾁｪｯｸﾎﾞｯｸｽ     +
'+---------------------------+
Private Sub Op_Sort_GotFocus(INDEX As Integer)
    For Each gObj In Op_Sort
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub Op_Sort_Click(INDEX As Integer)
    If ProcHB = "H1" Then Exit Sub
    Call DBInput
End Sub
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示     *
'*                                    *
'**************************************
Private Function DBInput(Optional InType As String = "") As Boolean
    Dim DateStr As String
    Dim DateSt  As String
    Dim DateEd  As String
    Dim BinSt   As String
    Dim BinEd   As String
    
    DBInput = True
    
    Call CisFrm.MousePT(11)
    
    On Error GoTo DBInput_Err:
' 抽出中メッセージ表示
    PNL_Wait.Visible = True
    PNL_PF.Enabled = False
    DoEvents
   
   
' 初期表示時、条件文作成(Where)
    If ProcHB = "H1" And InType = "" Then
        ''＜＜　検索条件作成　＞＞
        CisDBGrid.Select_Where = ""
        With CisFun
            .Where_CreateKB = NewDocument
            ' 年月度
            If Trim(iH1_YM) <> "" Then
                .Where_Create AlphaMe, "ZM.年月度", iH1_YM
            End If
            ' 契約先
            .Where_Create AlphaMe, "HM.得意先", iH1_Torcd
            ' 品番
            .Where_Create AlphaMe, "HM.品番", iH1_Hinbn
            ' 加工度
            Call CisComb1.CodeGet
            If Trim(CisComb1.Code) <> "" Then
               .Where_Create Numeric, "HM.品番区分", CisFun.Val2(CisComb1.Code), , Create
            End If
            CisDBGrid.Select_Where = .Where_Phrase
        End With
    End If
    Call Str_Monitor(CisDBGrid.Select_Where)


    '-- データ存在確認 -------
    gSL_Select = ""
    gSL_Select = gSL_Select & vbCrLf & "SELECT ZM.年月度, ZM.品番 入力品番, ZM.取引先,"
    gSL_Select = gSL_Select & vbCrLf & "       SUM(ISNULL(ZM.在庫数, 0)) 理論在庫数,"
    gSL_Select = gSL_Select & vbCrLf & "       CASE WHEN ISNULL(HM.表示品番, '')<> ''"
    gSL_Select = gSL_Select & vbCrLf & "            THEN HM.表示品番"
    gSL_Select = gSL_Select & vbCrLf & "            ELSE ISNULL(ZM.品番, '')"
    gSL_Select = gSL_Select & vbCrLf & "       END  品番,"
    gSL_Select = gSL_Select & vbCrLf & "       (SELECT ISNULL(MAX(背番号), '') 背番号 FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & vbCrLf & "         WHERE 品目 = 1 AND 品番 = ZM.品番 AND 取引先 = HM.得意先)"
    gSL_Select = gSL_Select & vbCrLf & "       AS 背番号,    "
    gSL_Select = gSL_Select & vbCrLf & "       HM.得意先 品番マスタ得意先,"
    gSL_Select = gSL_Select & vbCrLf & "       CASE WHEN HM.得意先 <> '' AND ISNULL(TR.略称, '') <> ''"
    gSL_Select = gSL_Select & vbCrLf & "           THEN HM.得意先 + ':' + ISNULL(TR.略称, '')"
    gSL_Select = gSL_Select & vbCrLf & "           ELSE CASE WHEN ISNULL(TR.略称, '') = ''"
    gSL_Select = gSL_Select & vbCrLf & "                     THEN HM.得意先"
    gSL_Select = gSL_Select & vbCrLf & "                     ELSE ISNULL(TR.略称, '')"
    gSL_Select = gSL_Select & vbCrLf & "                END"
    gSL_Select = gSL_Select & vbCrLf & "           END  契約先,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL(HM.品番区分, '') 品番区分,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL(MM.値名称  , '') 加工度,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL("
    gSL_Select = gSL_Select & vbCrLf & "           CASE HM.品番区分"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 1 /* 製品     */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN  (SELECT TOP(1) 単価 FROM 受注単価マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番   = ZM.品番"
    gSL_Select = gSL_Select & vbCrLf & "                               AND 取引先 = HM.得意先"
    gSL_Select = gSL_Select & vbCrLf & "                             ORDER BY 適用開始日付 DESC, 適用終了日付, 単価 DESC, 履歴NO DESC)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 2 /* 仕掛品   */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN /*HT.基準単価*/"
    gSL_Select = gSL_Select & vbCrLf & "                           (SELECT MAX(基準単価) 基準単価"
    gSL_Select = gSL_Select & vbCrLf & "                              FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番   = ZM.品番"
    gSL_Select = gSL_Select & vbCrLf & "                             /*AND 取引先 = HM.得意先*/)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 3 /* 購入品   */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN (SELECT TOP(1) 単価 FROM 購入単価マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番 = ZM.品番"
    gSL_Select = gSL_Select & vbCrLf & "                             ORDER BY 適用開始日付 DESC, 適用終了日付, 単価 DESC, 履歴NO DESC)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 5 /* 長期在庫 */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN 1"
    gSL_Select = gSL_Select & vbCrLf & "               ELSE 0 /* 4 無償品　または該当なしの場合 */"
    gSL_Select = gSL_Select & vbCrLf & "           END"
    gSL_Select = gSL_Select & vbCrLf & "       , 0) AS 単価,"
    gSL_Select = gSL_Select & vbCrLf & "       ROUND(ISNULL("
    gSL_Select = gSL_Select & vbCrLf & "           CASE HM.品番区分"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 1 /* 製品     */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN  (SELECT TOP(1) 単価 FROM 受注単価マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番   = ZM.品番"
    gSL_Select = gSL_Select & vbCrLf & "                               AND 取引先 = HM.得意先"
    gSL_Select = gSL_Select & vbCrLf & "                             ORDER BY 適用開始日付 DESC, 適用終了日付, 単価 DESC, 履歴NO DESC)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 2 /* 仕掛品   */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN /*HT.基準単価*/"
    gSL_Select = gSL_Select & vbCrLf & "                           (SELECT MAX(基準単価) 基準単価"
    gSL_Select = gSL_Select & vbCrLf & "                              FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番   = ZM.品番"
    gSL_Select = gSL_Select & vbCrLf & "                             /*AND 取引先 = HM.得意先*/)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 3 /* 購入品   */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN (SELECT TOP(1) 単価 FROM 購入単価マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番 = ZM.品番"
    gSL_Select = gSL_Select & vbCrLf & "                             ORDER BY 適用開始日付 DESC, 適用終了日付, 単価 DESC, 履歴NO DESC)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 5 /* 長期在庫 */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN 1"
    gSL_Select = gSL_Select & vbCrLf & "               ELSE 0 /* 4 無償品　または該当なしの場合 */"
    gSL_Select = gSL_Select & vbCrLf & "           END"
    gSL_Select = gSL_Select & vbCrLf & "       , 0) * SUM(ZM.在庫数), 0) 金額"
    gSL_Select = gSL_Select & vbCrLf & "  FROM 在庫マスタ ZM"
    gSL_Select = gSL_Select & vbCrLf & "  LEFT OUTER JOIN 品番マスタ HM"
    gSL_Select = gSL_Select & vbCrLf & "    ON HM.品番 = ZM.品番"
    gSL_Select = gSL_Select & vbCrLf & "  LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & vbCrLf & "    ON TR.取引先区分 = 0"
    gSL_Select = gSL_Select & vbCrLf & "   AND TR.取引先CD = HM.得意先"
    gSL_Select = gSL_Select & vbCrLf & "  LEFT OUTER JOIN 名称マスタ       MM"
    gSL_Select = gSL_Select & vbCrLf & "    ON MM.区分名称 = '品番区分'"
    gSL_Select = gSL_Select & vbCrLf & "   AND MM.値       = HM.品番区分"

    gSL_Select = gSL_Select & vbCrLf & CisDBGrid.Select_Where
    gSL_Select = gSL_Select & vbCrLf & " AND ZM.工順 = 1 "                              '--2020/02/10

    gSL_Select = gSL_Select & vbCrLf & " GROUP BY ZM.年月度, ZM.品番  , ZM.取引先, HM.表示品番,"
    gSL_Select = gSL_Select & vbCrLf & "          TR.略称  , HM.得意先, MM.値名称, HM.品番区分"
'-----------------------------------------------------------------------------------------------' 2009.12.22 add start
' 分類内容(Order By句) ｾｯﾄ
    With CisDBGrid
        .OrderSet (Op_Sort(0).Value)
        If .Select_OrderBy = "" Then
            DBInput = False
            GoTo DBInput_Ed
        End If
    End With
'    gSL_Select = gSL_Select & CisDBGrid.Select_Where & CisDBGrid.Select_OrderBy    ' 2019/07/16 DEL
'-----------------------------------------------------------------------------------------------' 2009.12.22 add end
    
    gSL_Select = gSL_Select & vbCrLf & CisDBGrid.Select_OrderBy                     ' 2019/07/16 ADD
    Call Str_Monitor(gSL_Select)            ' デバッグ用（外さなくても問題は無い）  ' 2019/07/16 ADD

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
    
    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
'===========================' 2019/07/16 DEL START
'    SendKeys "{END}"       '
'    SendKeys "{RIGHT}"
'    SendKeys "{LEFT}"
'    SendKeys "{HOME}"
'===========================' 2019/07/16 DEL END
    Exit Function
DBInput_Err:
    Call CisFrm.MousePT(0)
    CisFun.ErrorBox
    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TrueDBGrid)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+   グリッド内容初期設定
'+----------------------------------------------------------+
Private Sub GridInitSet()
' Call .GridInitSet(ﾏｰｼﾞ処理,分類処理,ﾃﾞｰﾀ連結名,印刷見出し,印刷桁数,合計印刷)
    With CisDBGrid
        Call .GridInitSet(, , "契約先", "契約先", 20)
        Call .GridInitSet(, , "品番", "品　　番", 30)
        Call .GridInitSet(, , "背番号", "背番号", 5)
        Call .GridInitSet(, , "理論在庫数", "理論在庫数", 20)
        Call .GridInitSet(, , "加工度", "加工度", 8)
        Call .GridInitSet(, , "単価", "単価", 7)
        Call .GridInitSet(, , "金額", "金額", 10)
    
        '=======================================================' 2019/07/09 DEL START
'         Call .GridInitSet(True, , "作業日", "作業日", 10)
'         Call .GridInitSet(True, , "部門", "部　門", 14)
'         Call .GridInitSet(True, , "作業者", "作業者", 14)
'         Call .GridInitSet(True, , "機械NO", "機械№", 10)
'         Call .GridInitSet(True, , "品番", "品　　番", 21)
'         Call .GridInitSet(True, , "工順", "工順", 4)
'         Call .GridInitSet(, , "加工数", "加工数", 10, True)
'         Call .GridInitSet(, , "開始", "開始", 5)
'         Call .GridInitSet(, , "終了", "終了", 5)
'         Call .GridInitSet(, , "不良数", "不良数", 8, True)
'         Call .GridInitSet(, , "段取", "段取", 3)
'         Call .GridInitSet(, , "作業日報管理NO", "手配NO", 10)
'         Call .GridInitSet(, , "基準工数", "基準工数", 8)
'         Call .GridInitSet(, , "目標工数", "目標工数", 8)
        '=======================================================' 2019/07/09 DEL END
    End With
End Sub
Private Sub GridInitSetCsv()
' Call .GridInitSet(ﾏｰｼﾞ処理,分類処理,ﾃﾞｰﾀ連結名,印刷見出し,印刷桁数,合計印刷)
    With CisDBGrid
        Call .GridInitSet(, , "契約先", "契約先", 20)
        Call .GridInitSet(, , "品番", "品　　番", 30)
        Call .GridInitSet(, , "背番号", "背番号", 5)
        Call .GridInitSet(, , "理論在庫数", "理論在庫数", 20)
        Call .GridInitSet(, , "加工度", "加工度", 8)
        Call .GridInitSet(, , "単価", "単価", 7)
        Call .GridInitSet(, , "金額", "金額", 10)
    
        '=======================================================' 2019/07/09 DEL START
'         Call .GridInitSet(True, , "作業日", "作業日", 10)
'         Call .GridInitSet(True, , "部門", "部　門", 30)
'         Call .GridInitSet(True, , "作業者", "作業者", 20)
'         Call .GridInitSet(True, , "機械NO", "機械№", 16)
'         Call .GridInitSet(True, , "品番", "品　　番", 26)
'         Call .GridInitSet(True, , "工順", "工順", 4)
'         Call .GridInitSet(, , "加工数", "加工数", 10, True)
'         Call .GridInitSet(, , "開始", "開始", 5)
'         Call .GridInitSet(, , "終了", "終了", 5)
'         Call .GridInitSet(, , "不良数", "不良数", 8, True)
'         Call .GridInitSet(, , "段取", "段取", 3)
'         Call .GridInitSet(, , "作業日報管理NO", "手配NO", 10)
'         Call .GridInitSet(, , "基準工数", "基準工数", 8)
'         Call .GridInitSet(, , "目標工数", "目標工数", 8)
        '=======================================================' 2019/07/09 DEL END
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
    Call CisDBGrid.MargeDisp(Op_Overlap(1).Value)
    Call KensuDisp

    With DBGrid1
         .SetFocus
    End With
    SendKeys "{HOME}"
    
    DBGridDisp = True
    
    On Error GoTo 0
    Exit Function
DBGridDisp_ER:
    If Err <> 40002 Then                        ' RDO タイムアウト
       CisFun.ErrorBox
    End If
End Function
'+-----------------------+
'+  グリッド内容初期化    +
'+-----------------------+
Private Sub DBGridClr()

    Call CisDBGrid.SplitsClr("", 0)             ' 分割ﾍﾟｲﾝ解除
    
    Adodc1.Recordset.Close
    DBGrid1.Close True
    
End Sub
'+----------------------------------------------------------+
'+ キー入力にて行移動 《 KeyUp イベント   》                  +
'+----------------------------------------------------------+
Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
        KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then
        Call KensuDisp
    Else
        Exit Sub
    End If
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 ColMove イベント   》            +
'+----------------------------------------------------------+
Private Sub DBGrid1_ColMove(ByVal Position As Integer, Cancel As Integer)
    CisDBGrid.ColMoveFlg = True
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 Paint イベント   》              +
'+----------------------------------------------------------+
Private Sub DBGrid1_Paint()
    If Not FormAct Then Exit Sub
    If CisDBGrid.ColMoveFlg Then                ' 再表示(分類)
        CisDBGrid.ColMoveFlg = False
        Call DBInput
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseDown  イベント   》               +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DBGrid1.RowContaining(Y) = -1 And _
        DBGrid1.ColContaining(X) = -1 Then
        Call CisDBGrid.SplitsClr("C", 0)        ' 分割ﾍﾟｲﾝのみ解除
        Call DBInput
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then        ' 分割ﾍﾟｲﾝ表示
        If Button = 2 Then
            Call CisDBGrid.SplitsAdd(DBGrid1.ColContaining(X))
            Call DBInput
        End If
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
      ''Call TotalDsp(DBGrid1.ColContaining(x)) ' 2019/07/16 DEL
    Else
        DBGrid1.Row = DBGrid1.RowContaining(Y)
    End If
    If DBGrid1.RowContaining(X) < 0 Then Exit Sub
    
    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
      ''Call TotalDsp(DBGrid1.ColContaining(x)) ' 2019/07/16 DEL
    Else
      ''If PNL_ExeMsg.Visible Then              ' 2019/07/16 DEL
      ''   PNL_ExeMsg.Visible = False           ' 2019/07/16 DEL
      ''End If                                  ' 2019/07/16 DEL
        DBGrid1.Row = DBGrid1.RowContaining(Y)
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call KensuDisp
'===================================' 2019/07/08 DEL START
'    If Button = 2 Then
'       PNL_ExeMsgX = ""
'       lb_TKSu = "": lb_TFSu = ""
'       PNL_ExeMsg.Visible = False
'    End If
'===================================' 2019/07/08 DEL END
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++    '++ デバッグ用です +++
'+                                                 +
'+  文字列をイミディエイトウィンドウに表示する     +
'+                                                 +
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Str_Monitor(ByVal wStr As String)
    Debug.Print vbCrLf
    Debug.Print Format(Now, "yyyy/MM/dd HH:mm:ss") + _
                " .................................................." + _
                vbCrLf + wStr
End Sub



