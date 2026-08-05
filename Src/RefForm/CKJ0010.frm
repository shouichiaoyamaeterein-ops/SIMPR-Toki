VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CKJ0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "品番マスタ検索"
   ClientHeight    =   8595
   ClientLeft      =   1395
   ClientTop       =   1350
   ClientWidth     =   13350
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
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
   ScaleHeight     =   8595
   ScaleWidth      =   13350
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1875
      Left            =   105
      ScaleHeight     =   1815
      ScaleWidth      =   8640
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   780
      Width           =   8700
      Begin Cis3D_v60.CIS3D UC_3D 
         Height          =   1725
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   3043
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
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   840
            Left            =   15
            Top             =   855
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "機 械 №"
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
            cPositionY      =   60
            Begin CisText_V60.CisText iH1_Kikai 
               Height          =   345
               Left            =   120
               TabIndex        =   3
               Top             =   360
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXX"
               MaxLength       =   3
            End
            Begin Cis3D_v60.CIS3D H1lb_Kikai 
               Height          =   330
               Left            =   690
               Top             =   360
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   582
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XX"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   840
            Index           =   5
            Left            =   2370
            Top             =   855
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "品        番"
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
            cPositionY      =   60
            Begin CisText_V60.CisText iH1_Hinbn 
               Height          =   375
               Left            =   105
               TabIndex        =   5
               Top             =   360
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
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   25
               IMEMode         =   2
               cDataReplace    =   1
               cFaZero         =   0
               cFbComma        =   0
            End
            Begin Cis3D_v60.CIS3D H1lb_Hinnm 
               Height          =   360
               Left            =   3270
               Top             =   360
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXXXXXXX2XX"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   840
            Left            =   15
            Top             =   15
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "作業日"
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
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisYMD_With_Btn.CisYMDwB iH1_Ymd 
               Height          =   345
               Left            =   90
               TabIndex        =   0
               Top             =   390
               Width           =   1890
               _ExtentX        =   3069
               _ExtentY        =   556
               cTextColor      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Object.Height          =   345
               Object.Width           =   1890
               cSize           =   1
               cChkResult      =   0   'False
               cTypeIntval     =   0
               BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Cis3D_v60.CIS3D Back_Torcd 
            Height          =   840
            Left            =   2115
            Top             =   15
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "部　　門"
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
            cPositionY      =   60
            Begin CisText_V60.CisText iH1_Torcd 
               Height          =   360
               Left            =   90
               TabIndex        =   1
               Top             =   390
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
            Begin Cis3D_v60.CIS3D H1lb_TorNm 
               Height          =   360
               Left            =   1080
               Top             =   390
               Width           =   2505
               _ExtentX        =   4419
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
         Begin Cis3D_v60.CIS3D Back_Sagyo 
            Height          =   840
            Left            =   5805
            Top             =   15
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "作 業 者"
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
            cPositionY      =   60
            Begin CisText_V60.CisText iH1_Sagyo 
               Height          =   360
               Left            =   90
               TabIndex        =   2
               Top             =   390
               Width           =   480
               _ExtentX        =   847
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
               Text            =   "XXX"
               MaxLength       =   3
            End
            Begin Cis3D_v60.CIS3D H1lb_Sagyo 
               Height          =   360
               Left            =   570
               Top             =   390
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXXXX"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   840
            Left            =   5805
            Top             =   855
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   1482
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
            cAlingnment     =   6
            cPositionY      =   60
         End
      End
   End
   Begin Cis3D_v60.CIS3D HLB_Title 
      Height          =   675
      Left            =   120
      Top             =   30
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   1191
      BackColor       =   12648384
      Caption         =   "手配№検索"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   180
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   150
      Width           =   855
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9630
      ScaleHeight     =   615
      ScaleWidth      =   3480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   30
      Width           =   3540
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   1755
         Top             =   15
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
         Left            =   2610
         Top             =   15
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
         BtnCaption      =   "戻る"
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   900
         Top             =   15
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   60
         Top             =   15
         Visible         =   0   'False
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
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5805
      Left            =   105
      ScaleHeight     =   5745
      ScaleWidth      =   13020
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2730
      Width           =   13080
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CKJ0010.frx":0000
         Height          =   5580
         Left            =   60
         OleObjectBlob   =   "CKJ0010.frx":0015
         TabIndex        =   4
         Top             =   60
         Width           =   12870
      End
   End
   Begin VB.PictureBox Back_Sort 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   7740
      ScaleHeight     =   570
      ScaleWidth      =   1710
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Width           =   1770
      Begin VB.OptionButton Op_Sort 
         BackColor       =   &H00808000&
         Caption         =   "降順"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   870
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
      Begin VB.OptionButton Op_Sort 
         BackColor       =   &H00808000&
         Caption         =   "昇順"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   105
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   105
      Top             =   150
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
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
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CKJ0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   作業日報テーブル検索
'**       フォームID    :   CKJ0010
'**       処理概要      :
'**
'**       作  成  日    :   2009/12/04  By CIS
'**       変  更  日    :   2011/08/11  By CIS SIMPR-A 検索画面,手配桁数幅拡張
'**       変  更  日    :   2014/02/04  By CIS WIN7対応
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim mSL_OrderA      As String       ' ソート順(Order BY)昇順
    Dim mSL_OrderD      As String       ' ソート順(Order BY)降順
    Dim mSL_SelectW     As String       ' SQL SELECT(Where)条件文
    Dim FormCap         As String
    
    Dim mNameTorcd      As String       ' 部門名
    Dim mNameSagyo      As String       ' 作業者名

    Dim CisFrm          As CISFormContorl
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    Dummy.Left = -1000
    Dummy.Enabled = False
'    Timer1.Interval = 1000   ' 時刻表示間隔
'    Timer1.Enabled = True    ' 時刻表示開始
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
''''    If RVI_Hinmoku = 9 Then
''''    Else
''''       H1Ck_Hinmoku(RVI_Hinmoku).Value = 1
''''    End If
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    If RV_Left <> 0 Or RV_Top <> 0 Then
        Me.Move RV_Left, RV_Top
    Else
        Call CisFrm.Frm_Center(False, x1024y768)
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormCap = Me.Caption
    Op_Sort(0).Value = True
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
    Call HeadBodyClear("H")
    RV_Rtn = False
    
    FormAct = Not FormAct
'
    If RV_Call <> "" Then
       Call ReturnPress
    End If
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
    Select Case Key_Code
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'   Me.Hide
'   If RV_Call = "Z" Then
'      Unload CKJ0010
'   Else
'      Me.Hide
'   End If
    Unload CKJ0010
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
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        PB_Ent.BtnCaption = "開始"
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Ymd" Then iH1_Ymd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Sagyo" Then Call Look_Sagyo: Exit Sub
    If PB_Look.Tag = "iH1_Kikai" Then Call Look_Kikai: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+----------------------------+
'+       部門検索             +
'+----------------------------+
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
'+----------------------------+
'+       作業者検索           +
'+----------------------------+
Private Sub Look_Sagyo()
    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "担当管理"
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_Sagyo = RV_Code          ' 値
       H1lb_Sagyo = RV_Kubun        ' 値名称
    End If
End Sub
'+----------------------------+
'+       機械№検索           +
'+----------------------------+
Private Sub Look_Kikai()
    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "機械NO"
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_Kikai = RV_Code          ' 値
       H1lb_Kikai = RV_Kubun        ' 値名称
    End If
End Sub
'+----------------------------+
'+       品番検索             +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
   'RV_Call = "C"
    RVI_Hinmoku = 9
    
   'CKK0030.Show vbModal        ' 品番検索時は0030
   'Unload CKK0030
   'Set CKK0030 = Nothing
    RVI_Torcd = Trim(iH1_Torcd)
    CKK0035.Show vbModal        ' 品番取引先検索時は0035
    Unload CKK0035
    Set CKK0035 = Nothing
    iH1_Hinbn = RV_Hinbn
    H1lb_Hinnm = RV_Hinnm
End Sub
'+-----------------------------+
'+  作業日
'+-----------------------------+
Private Sub iH1_Ymd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  部門
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
    
    H1lb_TorNm = ""
    If Trim(iH1_Torcd) <> "" Then
        gSL_Select = "SELECT * FROM 取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先CD = '" & Trim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 社内区分 = 1"
        If TRMRead(gSL_Select) Then
            H1lb_TorNm = TRM.正式名称
        Else
            H1lb_TorNm = ""
        End If
    End If
End Sub
'+-----------------------------+
'+  作業者
'+-----------------------------+
Private Sub iH1_Sagyo_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Sagyo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    H1lb_Sagyo = ""
    If Trim(iH1_Sagyo) <> "" Then
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '担当管理'"
        gSL_Select = gSL_Select & "   AND 値       = '" & Trim(iH1_Sagyo) & "'"
        If SYMRead(gSL_Select) Then
            H1lb_Sagyo = SYM.値名称
        Else
            H1lb_Sagyo = ""
        End If
    End If
End Sub
'+-----------------------------+
'+  機械
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
    
    H1lb_Kikai = ""
    If Trim(iH1_Kikai) <> "" Then
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '機械NO'"                  ' 2010.01.06 upd
        gSL_Select = gSL_Select & "   AND 値       = '" & Trim(iH1_Kikai) & "'"
        If SYMRead(gSL_Select) Then
            H1lb_Kikai = SYM.値名称
        Else
            H1lb_Kikai = ""
        End If
    End If
End Sub
'+-----------------------------+
'+  品番
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

    If Trim(iH1_Hinbn) <> "" Then
        gSL_Select = "SELECT * FROM 品番マスタ"
        gSL_Select = gSL_Select & " WHERE 品番 = '" & Trim(iH1_Hinbn) & "'"
        If HNMRead(gSL_Select) Then
            H1lb_Hinnm = HNM.品名
        Else
            H1lb_Hinnm = ""
        End If
    End If
End Sub
''+----------------------------+
''+       手配先検索           +
''+----------------------------+
'Private Sub Look_Torcd()
'    RV_Left = 0
'    RV_Top = 0
'    If H1Ck_Hinmoku(0).Value <> H1Ck_Hinmoku(1).Value Then
'       RV_TorKb = H1Ck_Hinmoku(1).Value
'    Else
'       RV_TorKb = 2
'    End If
'
'    RV_Call = ""
'
'    CKK0020.Show vbModal
'    Unload CKK0020
'    Set CKK0020 = Nothing
'
'    If RV_Rtn Then
'        iH1_Torcd = RV_TorcdK
'        H1lb_Tornm = RV_TorRName
'    End If
'End Sub
''+-----------------------------+
''+  iH1_Torcd
''+-----------------------------+
'Private Sub iH1_Torcd_GotFocus()
'    PB_Look.Tag = ActiveControl.Name
'    PB_Look.Visible = True
'End Sub
'Private Sub iH1_Torcd_LostFocus()
'    If ActiveControl.Name <> "PB_Look" Then
'        PB_Look.Tag = ""
'        PB_Look.Visible = False
'    End If
'    H1lb_Tornm = ""
'    If TorNmGet(iH1_Torcd, 0) Then
'       H1lb_Tornm = TRM_RName
'    End If
'End Sub
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
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1"
          H1Mode = True
          H1Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_Look.Visible = False
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
    If HeadBody = "B" Then
       DBGrid1.Close True
       Adodc1.Recordset.Close
       Adodc1.RecordSource = ""
       Me.Caption = FormCap
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
       If Not Head1Chk Then GoTo ReturnPress_Ed
       If Not BodySet Then GoTo ReturnPress_Ed
       ProcHB = "B1"
       PB_Ent.BtnCaption = "選択"
       Call DispChange(ProcHB)
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then Call RtnSet
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    If Not (iH1_Ymd = "" Or Len(Trim(iH1_Ymd)) = 8) Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(1) = "  年月日を入力して下さい    "
            .MB_Button = OK
            .MBOX
        End With
'       ReDim gMBMsg(5)
'       gMBMsg(2) = "  年月日を入力して下さい  "
'       gMBMsg(3) = "   "
'       gMBTitle = " "
'       gMBButton = "E"
'       gRtn = MBOX
       iH1_Ymd.SetFocus
       Exit Function
    End If
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
'    If iB1_???? = "" Then
'       ReDim gMBMsg(5)
'       gMBMsg(2) = "  ???を入力して下さい  "
'       gMBMsg(3) = "   "
'       gMBTitle = " "
'       gMBButton = "E"
'       gRtn = MBox
'       iB1_???.SetFocus
'       Exit Function
'    End If
    Body1Chk = True
End Function
'****************************
'*      分類切替
'****************************
Private Sub Op_Sort_Click(INDEX As Integer)
    If ProcHB = "H1" Then Exit Sub
    If Adodc1.Recordset.RecordCount <> 0 Then Call BodySet
    SendKeys "{TAB}"
End Sub
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet() As Boolean
    
    BodySet = False
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT SN.作業日報管理NO, SN.受入,"                                                                      ' 2010.01.06 upd : 受入を追加 / 作業日を削除
    gSL_Select = gSL_Select & "       SUBSTRING(SN.作業日,1,4)+'/'+SUBSTRING(SN.作業日,5,2)+'/'+SUBSTRING(SN.作業日,7,2) AS 作業日, "   ' 2010.01.06 add
    gSL_Select = gSL_Select & "                 ISNULL(SN.部門,  '')                + ':' + TM.正式名称 AS 部門  ,"
    gSL_Select = gSL_Select & "       SUBSTRING(ISNULL(SN.作業者,'')+'   '  , 1, 3) + ':' + M2.値名称   AS 作業者,"
    gSL_Select = gSL_Select & "       SUBSTRING(ISNULL(SN.機械NO,'')+'   '  , 1, 3) + ':' + M1.値名称   AS 機械NO,"
   'gSL_Select = gSL_Select & "       SUBSTRING(ISNULL(SN.品番,  '')+'                         ', 1,25) + ':'"
   'gSL_Select = gSL_Select & "               + ISNULL(HM.品名,  '') AS 品番,"
    gSL_Select = gSL_Select & "                 ISNULL(SN.品番,  '') AS 品番,"
    gSL_Select = gSL_Select & "                 ISNULL(HM.表示品番, '') AS 表示品番,"               ' 2010.01.06 add
    gSL_Select = gSL_Select & "       ISNULL(SN.加工数,0)  AS 加工数,"
    gSL_Select = gSL_Select & "       ISNULL(M1.値名称,'') AS 機械NO名, ISNULL(M2.値名称,'')   AS 作業者名,"
    gSL_Select = gSL_Select & "       ISNULL(HM.品名  ,'') AS 品名 ,    ISNULL(TM.正式名称,'') AS 部門名   "
    gSL_Select = gSL_Select & " FROM  作業日報テーブル SN"
    
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ M1"
    gSL_Select = gSL_Select & "            ON  M1.区分名称 = '機械NO'"                              ' 2010.01.06 upd
    gSL_Select = gSL_Select & "            AND M1.値       = SN.機械NO"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ M2"
    gSL_Select = gSL_Select & "            ON  M2.区分名称 = '担当管理'"
    gSL_Select = gSL_Select & "            AND M2.値       = SN.作業者"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番マスタ HM"
    gSL_Select = gSL_Select & "            ON  HM.品番     = SN.品番"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM"
    gSL_Select = gSL_Select & "            ON  TM.取引先CD = SN.部門"
    gSL_Select = gSL_Select & "            AND TM.社内区分 = 1"
     
    mSL_OrderA = " Order By SN.作業日報管理NO, SN.作業日 "
    mSL_OrderD = " Order By SN.作業日報管理NO DESC, SN.作業日"

' 選択条件セット
    mSL_SelectW = ""
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create AlphaMe, "SN.作業日", iH1_Ymd
        .Where_Create AlphaMe, "SN.部門", iH1_Torcd
        .Where_Create AlphaMe, "SN.作業者", iH1_Sagyo
        .Where_Create AlphaMe, "SN.機械NO", iH1_Kikai
        .Where_Create AlphaMe, "SN.品番", iH1_Hinbn
        
        mSL_SelectW = .Where_Phrase
    End With
    
    mSL_SelectW = mSL_SelectW & " GROUP BY SN.作業日報管理NO, SN.作業日, SN.部門, SN.作業者, SN.機械NO, SN.品番, SN.加工数,"
    mSL_SelectW = mSL_SelectW & "          M1.値名称, M2.値名称, HM.品名, TM.正式名称, SN.受入, HM.表示品番"    ' 2010.01.06 upd : 受入と表示品番を追加
    
    If Op_Sort(0) Then
        Adodc1.RecordSource = gSL_Select & mSL_SelectW & mSL_OrderA
    Else
        Adodc1.RecordSource = gSL_Select & mSL_SelectW & mSL_OrderD
    End If
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount = 0 Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(1) = "    該当するデータはありません  "
            .MB_Button = OK
            .MBOX
        End With
        Exit Function
    End If
    Call TotalDisp
    BodySet = True
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TDBGridPro)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+      ADODC
'+----------------------------------------------------------+
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Adodc1.RecordSource <> "" Then Call TotalDisp
End Sub
'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub TotalDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    Me.Caption = FormCap & Space(5)
    Me.Caption = Me.Caption & "抽出件数 = " & Format(NowCnt, "##,##0")
    Me.Caption = Me.Caption & "/" & Format(MaxCnt, "##,##0") & " 件"
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SendKeys "{RIGHT}"
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If DBGrid1.ColContaining(X) < 0 Then Exit Sub
    DBGrid1.Row = DBGrid1.RowContaining(Y)
    DoEvents
    Call RtnSet
End Sub
'+-------------------------------------+
'+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
'+-------------------------------------+
Private Sub RtnSet()
    Dim wPos As Integer
    
    With Adodc1                                             '==== 格納先は品番のデータを使用 =====
        RV_Tehai = .Recordset("作業日報管理NO")                     ' 手配№
        RV_Hinbn = Trim(.Recordset("品番"))                         ' 品番
        RV_Ukeir = .Recordset("作業日")                             ' 年月日  ：受入で代用
        
        wPos = InStr(1, .Recordset("部門"), ":")
        If wPos <> 0 Then
           RV_Torcd = Trim(Mid(.Recordset("部門"), 1, wPos - 1))    ' 部門
           RVI_Torcd = Trim(Mid(.Recordset("部門"), wPos + 1))      ' 部門名
        End If
        
        RV_Seban = Trim(Mid(.Recordset("作業者"), 1, 3))            ' 作業者  ：背番号で代用
        RVI_Seban = Trim(Mid(.Recordset("作業者"), 5))              ' 作業者名：背番で代用
        RVI_Ukeir = Trim(.Recordset("受入"))                        ' 受入                  ' 2010.01.06 add
    End With
    Adodc1.Recordset.Close
    RV_Rtn = True
    Call PB_END_Click
End Sub
