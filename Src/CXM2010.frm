VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form CXM2010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "î≠íçÇ©ÇÒÇŒÇÒêFÉRÅ[Éhê›íË"
   ClientHeight    =   8550
   ClientLeft      =   5055
   ClientTop       =   2715
   ClientWidth     =   6450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   6450
   Begin VB.PictureBox H1_Area2 
      Height          =   3495
      Left            =   150
      ScaleHeight     =   3435
      ScaleWidth      =   6075
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4080
      Width           =   6135
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   3375
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   5985
         _Version        =   196608
         _ExtentX        =   10557
         _ExtentY        =   5953
         _StockProps     =   228
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "áÇ|éÛì¸|êFNo|îwåi           |êFNo.|ògêF          |"
         Rows            =   11
         Cols            =   7
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         RowHeightMin    =   300
         SelectionMode   =   1
         Begin Cis3D_v60.CIS3D Input_Area 
            Height          =   285
            Left            =   330
            Top             =   300
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   503
            BackColor       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin VB.TextBox B_TabDummy 
               Height          =   300
               Left            =   5610
               TabIndex        =   43
               Text            =   "Text1"
               Top             =   60
               Width           =   195
            End
            Begin CisText_V60.CisText iH1_Ukeir 
               Height          =   285
               Left            =   0
               TabIndex        =   0
               Top             =   0
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XX"
               MaxLength       =   2
               cDataReplace    =   1
               cFbComma        =   0
            End
            Begin Cis3D_v60.CIS3D B1lb_Waku 
               Height          =   315
               Left            =   3600
               Top             =   0
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               BackColor       =   16777215
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin CisText_V60.CisText iH1_Waku 
               Height          =   285
               Left            =   3015
               TabIndex        =   2
               Top             =   0
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   503
               cFormat         =   "##"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
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
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
            Begin Cis3D_v60.CIS3D B1lb_Haikei 
               Height          =   315
               Left            =   1140
               Top             =   0
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               BackColor       =   16777215
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin CisText_V60.CisText iH1_Haikei 
               Height          =   285
               Left            =   570
               TabIndex        =   1
               Top             =   0
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               cFormat         =   "##"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
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
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   2595
      Left            =   150
      ScaleHeight     =   2535
      ScaleWidth      =   6075
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1410
      Width           =   6135
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   495
         Left            =   360
         Top             =   0
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   873
         ForeColor       =   128
         Caption         =   "êFÅ@ê›Å@íË"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   0
         Left            =   360
         Top             =   495
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   1
            Left            =   1290
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   1
            Left            =   60
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            ForeColor       =   0
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'é¿ê¸
            Height          =   300
            Index           =   1
            Left            =   930
            TabIndex        =   13
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   495
         Left            =   30
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   873
         ForeColor       =   128
         Caption         =   "êFáÇ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cPositionX      =   40
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   405
         Left            =   30
         Top             =   495
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "ÇP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   1
         Left            =   2370
         Top             =   495
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   6
            Left            =   1290
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   6
            Left            =   60
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   6
            Left            =   930
            TabIndex        =   15
            Top             =   30
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   405
         Left            =   2040
         Top             =   495
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "ÇU"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   3
         Left            =   360
         Top             =   900
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   2
            Left            =   1290
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   2
            Left            =   60
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   2
            Left            =   930
            TabIndex        =   17
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   405
         Left            =   30
         Top             =   900
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "ÇQ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   4
         Left            =   2370
         Top             =   900
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   7
            Left            =   1320
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   7
            Left            =   90
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   7
            Left            =   960
            TabIndex        =   19
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   405
         Left            =   2040
         Top             =   900
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "ÇV"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   5
         Left            =   360
         Top             =   1305
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   3
            Left            =   1290
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   3
            Left            =   60
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   3
            Left            =   930
            TabIndex        =   21
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   405
         Left            =   30
         Top             =   1305
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "ÇR"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   6
         Left            =   2370
         Top             =   1305
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   8
            Left            =   1320
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   8
            Left            =   90
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   8
            Left            =   960
            TabIndex        =   23
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   405
         Left            =   2040
         Top             =   1305
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "ÇW"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   7
         Left            =   360
         Top             =   1710
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   4
            Left            =   1290
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   4
            Left            =   60
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   4
            Left            =   930
            TabIndex        =   25
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   405
         Left            =   30
         Top             =   1710
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "ÇS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   8
         Left            =   2370
         Top             =   1710
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   9
            Left            =   1320
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   9
            Left            =   90
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   9
            Left            =   960
            TabIndex        =   27
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   405
         Left            =   2040
         Top             =   1710
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "ÇX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   9
         Left            =   360
         Top             =   2115
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   5
            Left            =   1290
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   5
            Left            =   60
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   5
            Left            =   930
            TabIndex        =   29
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   405
         Left            =   30
         Top             =   2115
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "ÇT"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   10
         Left            =   2370
         Top             =   2115
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   10
            Left            =   1320
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   10
            Left            =   90
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   10
            Left            =   960
            TabIndex        =   31
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   405
         Left            =   2040
         Top             =   2115
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "10"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   495
         Left            =   2370
         Top             =   0
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   873
         ForeColor       =   128
         Caption         =   "êFÅ@ê›Å@íË"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   495
         Left            =   2040
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   873
         ForeColor       =   128
         Caption         =   "êFáÇ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cPositionX      =   40
      End
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   495
         Index           =   0
         Left            =   4380
         Top             =   0
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   873
         ForeColor       =   128
         Caption         =   "êFÅ@ê›Å@íË"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   495
         Index           =   0
         Left            =   4050
         Top             =   0
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   873
         ForeColor       =   128
         Caption         =   "êFáÇ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cPositionX      =   40
      End
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   405
         Left            =   4050
         Top             =   495
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "11"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   405
         Left            =   4050
         Top             =   900
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "12"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D20 
         Height          =   405
         Left            =   4050
         Top             =   1305
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "13"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   14
         Left            =   4380
         Top             =   1710
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   14
            Left            =   1320
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   14
            Left            =   90
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   14
            Left            =   960
            TabIndex        =   33
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D21 
         Height          =   405
         Left            =   4050
         Top             =   1710
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "14"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D22 
         Height          =   405
         Left            =   4050
         Top             =   2115
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   714
         BackColor       =   14737632
         Caption         =   "15"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   11
         Left            =   4380
         Top             =   495
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   11
            Left            =   1320
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   11
            Left            =   90
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   11
            Left            =   960
            TabIndex        =   35
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   12
         Left            =   4380
         Top             =   900
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   12
            Left            =   1320
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   12
            Left            =   90
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   12
            Left            =   960
            TabIndex        =   37
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   13
         Left            =   4380
         Top             =   1305
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   13
            Left            =   1320
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   13
            Left            =   90
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            ForeColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   13
            Left            =   960
            TabIndex        =   39
            Top             =   35
            Width           =   330
         End
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   15
         Left            =   4380
         Top             =   2115
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   285
            Index           =   15
            Left            =   1320
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   60
            Width           =   315
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   300
            Index           =   15
            Left            =   90
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   300
            Index           =   15
            Left            =   960
            TabIndex        =   41
            Top             =   35
            Width           =   330
         End
      End
   End
   Begin VB.PictureBox H1_Area0 
      Height          =   855
      Left            =   150
      ScaleHeight     =   795
      ScaleWidth      =   6075
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   6135
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   810
         Left            =   0
         Top             =   0
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   1429
         ForeColor       =   128
         Caption         =   "ïWèÄ(ñ¢ê›íËéÛì¸)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   2
         Left            =   2160
         Top             =   0
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   714
         ForeColor       =   128
         Caption         =   "  îwåiêF"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   3390
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   60
            Width           =   375
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   330
            Index           =   0
            Left            =   1050
            Top             =   30
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   582
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H80000008&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   330
            Index           =   0
            Left            =   2820
            TabIndex        =   8
            Top             =   45
            Width           =   540
         End
      End
      Begin Cis3D_v60.CIS3D Back_Iro 
         Height          =   405
         Index           =   16
         Left            =   2160
         Top             =   405
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   714
         ForeColor       =   128
         Caption         =   "    ògêF"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         Begin VB.CommandButton PB_Color 
            Caption         =   "..."
            Height          =   315
            Index           =   16
            Left            =   3390
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   60
            Width           =   375
         End
         Begin Cis3D_v60.CIS3D H1lb_SmpChr 
            Height          =   330
            Index           =   16
            Left            =   1050
            Top             =   30
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   582
            BackColor       =   16777215
            Caption         =   "SAMPLE"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   14737632
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin VB.Label H1lb_SmpClr 
            BackColor       =   &H80000008&
            BorderStyle     =   1  'é¿ê¸
            Caption         =   "Label1"
            Height          =   330
            Index           =   16
            Left            =   2820
            TabIndex        =   10
            Top             =   45
            Width           =   540
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -420
      Top             =   525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'µÃå≈íË
      Left            =   -135
      TabIndex        =   4
      Text            =   "Dummy"
      Top             =   450
      Width           =   180
   End
   Begin Cis3D_v60.CIS3D PB_ENT 
      Height          =   645
      Left            =   1170
      Top             =   7740
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1138
      ForeColor       =   128
      Caption         =   "ì¸óÕ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   645
      Left            =   3660
      Top             =   7740
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1138
      ForeColor       =   12582912
      Caption         =   "èIóπ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "Åy î≠íçÇ©ÇÒÇŒÇÒêFÉRÅ[Éhê›íË  Åz"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   9540
         Top             =   60
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
   End
   Begin VB.Label BaseColor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   345
      Left            =   270
      TabIndex        =   5
      Tag             =   "6008319"
      Top             =   8790
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "CXM2010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   Ç©ÇÒÇŒÇÒêFÉRÅ[Éhê›íË
'**       ÉtÉHÅ[ÉÄID    :   CXM2010
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2008/06/09  By CIS
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB                  As String * 2   ' ì¸óÕàÊêÿë÷
    
    Dim CisVsGrid               As New CisVsGrid3
    
    Dim mKubun_Iro              As String
    Dim mKubun_Ukeir            As String
    Dim mColorMax               As Long
    Dim mSaveRow                As Long
    Dim wKeyCode                As Integer
    
    Dim wRow                    As Integer
    Dim ActObj                  As Object

'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt            +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # ìÒ èd ãN ìÆ ñh é~  #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # èâ ä˙ ì‡ óe éÊ ìæ  #
'   #-------------------#
    If Not IniGet Then End
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
'   #-------------------#
'   #  çÄ ñ⁄ èâ ä˙ âª    #
'   #-------------------#
    Call CisFrm.InitFld
    
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)

'   #------------------#
'   #  ∏ﬁÿØƒﬁèâä˙èàóù   #
'   #------------------#
    With CisVsGrid
        Set .GridObj = vsGrid1              ' ∏ﬁÿØƒﬁµÃﬁºﬁ™∏ƒñº æØƒ
        .FixedHeight = 300                  ' å≈íËçs(å©èoÇµ)çÇÇ≥ê›íË(è»ó™éûÇÕ300)
        .PatanMax = 1                       ' ⁄≤±≥ƒ ﬂ¿∞›ç≈ëÂêî
        .PatanInit = 0                      ' ⁄≤±≥ƒ ﬂ¿∞›èâä˙íl
        .SelectCol = 10                     ' ÉZÉãëIëèàóù‹∞∏óÒ
        .NumberSet = All
'        .EditCheckCol = 1
'        Set .SelectDispObj = B1lb_SKensu    ' ÉZÉãëIëåèêîï\é¶µÃﬁºﬁ™∏ƒæØƒ
' Ãﬂ€∏ﬁ◊—çÏê¨éûÅAà»â∫ÇÃÃ◊∏ﬁÇæØƒÇµÅA∏ÿØÃﬂŒﬁ∞ƒﬁÇÊÇË∫∞ƒﬁÇÃì\ÇËïtÇØÇçsÇ§ÅB(GridInit)
        .InitGet = False                    ' ∏ﬁÿØƒﬁì‡óeèâä˙ê›íË ∏ÿØÃﬂŒﬁ∞ƒﬁçÏê¨Ã◊∏ﬁ

        .Init                               ' ∏ﬁÿØƒﬁèâä˙èàóù é¿çs
    End With
    Call GridInit                           ' ∏ﬁÿØƒﬁèâä˙ï\é¶
    
'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(False)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct
    mColorMax = 15
    
    Set ActObj = iH1_Ukeir
    
    Call InitColorDisp
    Call GridSet
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås ÇjÇÖÇôÇcÇèÇóÇéÉCÉxÉìÉg Åt                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄÅs ÉtÉ@ÉìÉNÉVÉáÉìÉLÅ[ópÉAÉNÉZÉâÉåÅ[É^ Åt         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    
    Select Case Key_Code
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      'Åyì¸óÕÅz
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function

'+---------------------+
'+    èI  óπ  ÉL Å[     +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect                            ' √ﬁ∞¿Õﬁ∞Ωêÿíf
    Call CisFrm.UnLoadDisp                          ' Ã´∞—µÃﬁºﬁ™∏ƒè¡ãé
    End
End Sub
'+----------------------------+
'+    ì¸  óÕ  ÉL Å[(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'+------------------+
'+    êFÉ_ÉCÉAÉçÉO
'+------------------+
Private Sub PB_Color_Click(Index As Integer)
    Dim SaveRow         As Long
    With CommonDialog1
        .CancelError = True
        On Error GoTo ErrHandler
        
        .Color = H1lb_SmpClr(Index).BackColor
        .Flags = cdlCCRGBInit Or cdlCCPreventFullOpen
        .Flags = cdlCCRGBInit
        .ShowColor
                
        H1lb_SmpClr(Index).BackColor = .Color
        H1lb_SmpChr(Index).ForeColor = .Color
        If .Color <> BaseColor.BackColor Then
           H1lb_SmpChr(Index) = "SAMPLE"
        Else
            If Index <> 0 And Index <> 16 Then
               H1lb_SmpChr(Index) = ""
            End If
        End If
    End With

    With vsGrid1
        .Redraw = False
        SaveRow = .Row
        For gCnt1 = 1 To .Rows - 1
            If CisFun.Val2(.TextMatrix(gCnt1, 2)) = Index Then
                .Row = gCnt1
                Call GridColorSet(True, CLng(Index), 3)
            End If
            If CisFun.Val2(.TextMatrix(gCnt1, 4)) = Index Then
                .Row = gCnt1
                Call GridColorSet(True, CLng(Index), 5)
            End If
        Next gCnt1
        .Row = SaveRow: .Col = 1
        .Redraw = True
        If CisFun.Val2(iH1_Haikei) = Index Then
           B1lb_Haikei.BackColor = H1lb_SmpClr(Index).BackColor
        End If
        If CisFun.Val2(iH1_Waku) = Index Then
           B1lb_Waku.BackColor = H1lb_SmpClr(Index).BackColor
        End If
    End With

ErrHandler:
    On Error GoTo 0

'    iH1_Iro(Index).SetFocus
End Sub
'+----------------------+
'+  iH1_Ukeir
'+----------------------+
Private Sub iH1_Ukeir_GotFocus()
    Set ActObj = iH1_Ukeir
    B_TabDummy.Tag = "S"
End Sub
Private Sub iH1_Ukeir_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown And KeyCode <> 0 Then Exit Sub
       KeyCode = wKeyCode
       Dummy.Enabled = True: Dummy.SetFocus
       iH1_Ukeir.SetFocus
       Dummy.Enabled = False
       KeyCode = 0
End Sub
Private Sub iH1_Ukeir_LostFocus()
    vsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 1) = iH1_Ukeir
End Sub
'+----------------------+
'+  iH1_Haikei
'+----------------------+
Private Sub iH1_Haikei_GotFocus()
    Set ActObj = iH1_Haikei
End Sub
Private Sub iH1_Haikei_KeyDown(KeyCode As Integer, Shift As Integer)
    wKeyCode = KeyCode
    Call Grid_RowChange(KeyCode)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown And KeyCode <> 0 Then Exit Sub
       KeyCode = wKeyCode
       Dummy.Enabled = True: Dummy.SetFocus
       iH1_Haikei.SetFocus
       Dummy.Enabled = False
       KeyCode = 0
    
End Sub
Private Sub iH1_Haikei_LostFocus()
    vsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 2) = iH1_Haikei
    If CisFun.Val2(iH1_Haikei) > 0 And CisFun.Val2(iH1_Haikei) < 16 Then
       B1lb_Haikei.BackColor = H1lb_SmpClr(CisFun.Val2(iH1_Haikei)).BackColor
    Else
       B1lb_Haikei.BackColor = H1lb_SmpClr(0).BackColor
    End If
End Sub
'+----------------------+
'+  iH1_Waku
'+----------------------+
Private Sub iH1_Waku_GotFocus()
    Set ActObj = iH1_Waku
    B_TabDummy.Tag = "E"
End Sub
Private Sub iH1_Waku_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown And KeyCode <> 0 Then Exit Sub
       KeyCode = wKeyCode
       Dummy.Enabled = True: Dummy.SetFocus
       iH1_Waku.SetFocus
       Dummy.Enabled = False
       KeyCode = 0
End Sub
Private Sub iH1_Waku_LostFocus()
    vsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 4) = iH1_Waku
    If CisFun.Val2(iH1_Waku) > 0 And CisFun.Val2(iH1_Waku) < 16 Then
       B1lb_Waku.BackColor = H1lb_SmpClr(CisFun.Val2(iH1_Waku)).BackColor
    Else
       B1lb_Waku.BackColor = H1lb_SmpClr(mColorMax + 1).BackColor
    End If
End Sub

'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If B_TabDummy.Tag = "S" Then
        Call TextToGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If B_TabDummy.Tag = "E" Then
        Call TextToGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    SendKeys "{Tab}"
End Sub
'+------------------------+
'+    ÉZÉãÇÃîzêF          +
'+------------------------+
Private Sub GridColorSet(Optional SetFlg As Boolean = False, Optional ColorNo As Long = 0, Optional TCol As Long = 0, Optional KB As Integer = 0)
    Dim cRow    As Integer
    
    If Not SetFlg Then
        ColorNo = CisFun.Val2(iH1_Haikei)
    End If
    With vsGrid1
        If KB <> 0 Then
            cRow = .Row
            .Row = CisFun.Val2(Input_Area.Tag)
        End If
        For gCnt3 = TCol To TCol
            .Col = gCnt3
            If Trim(.TextMatrix(.Row, TCol - 1)) <> "" Then
                If TCol = 2 Then
                   .CellBackColor = H1lb_SmpClr(0).BackColor
                Else
                   .CellBackColor = H1lb_SmpClr(16).BackColor
                End If
            Else
                .CellBackColor = 0
                Exit For
            End If
            If ColorNo <= mColorMax Then
               If Trim(H1lb_SmpChr(ColorNo).Caption) <> "" Then
                  .CellBackColor = H1lb_SmpClr(ColorNo).BackColor
               End If
            End If
        Next gCnt3
       .Col = 1
       If KB <> 0 Then
          .Row = cRow
       End If
    End With
    
 End Sub
'****************************
'*                          *
'*      HEAD/BODY èàóù      *
'*                          *
'****************************
'****************************
'*      ì¸óÕêÿë÷ èàóù        *
'****************************
Private Function DispChange(Pro As String)
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    
    H1Mode = False:     H1Color = gPLostSel
    
'   ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
    End Select
    
    H1_Area0.Enabled = H1Mode
    H1_Area1.Enabled = H1Mode
    H1_Area2.Enabled = H1Mode
    H1_Area0.BackColor = H1Color
    H1_Area1.BackColor = H1Color
    H1_Area2.BackColor = H1Color
    
    PB_END.Visible = H1Mode
    
'   ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
    Dummy.Enabled = False
End Function
'****************************
'*      ì¸óÕèàóù(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        Call TextToGrid(Val(Input_Area.Tag))
        Call GridTextEnabled(False)
        If Not Head1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
   Call vsGrid1_Click
End Sub
'**************************************
'*      Head1¡™Ø∏èàóù
'**************************************
Private Function Head1Chk() As Boolean
    Dim wHaikei         As Integer
    Dim wWaku           As Integer
    Dim Cnt2            As Long
    Dim wUkeir          As String
    
    Head1Chk = False
'
    mSaveRow = vsGrid1.Row
' çÏã∆É`ÉFÉbÉN
    With vsGrid1
        
        For gInt = 1 To .Rows - 1
        
            ' éÛì¸èdï°¡™Ø∏
            wUkeir = RTrim(.TextMatrix(gInt, 1))
            For Cnt2 = gInt + 1 To .Rows - 1
                If RTrim(.TextMatrix(Cnt2, 1)) <> "" Then
                    If wUkeir = RTrim(.TextMatrix(Cnt2, 1)) Then
                        With CisFun
                             .MB_Lines = 4
                             .MB_MSG(1) = "    ======== " & gInt & "çsñ⁄ éÛì¸ =========    "
                             .MB_MSG(3) = "    Å@Å@" & Cnt2 & "çsñ⁄Ç∆èdï°ÇµÇƒÇ¢Ç‹Ç∑ÅB           "
                             .MB_Title = ""
                             .MB_Button = Error
                             .MBOX
                        End With
                        .Row = gInt
                        Input_Area.Tag = gInt
                        Call GridToText(Input_Area.Tag)
                        Set ActObj = iH1_Ukeir
                        Exit Function
                    End If
                End If
            Next Cnt2
            
            wHaikei = CisFun.Val2(.TextMatrix(gInt, 2))
            wWaku = CisFun.Val2(.TextMatrix(gInt, 4))
            ' îwåiêFáÇ
            If wHaikei > 15 And wHaikei <> 0 Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(1) = "    ======== " & gInt & "çsñ⁄ îwåiêF =========    "
                     .MB_MSG(3) = "    êFáÇÇÕ ÇPÅ`ÇPÇTÇÃä‘ÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢         "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                .Row = gInt
                Input_Area.Tag = gInt
                Call GridToText(Input_Area.Tag)
                Set ActObj = iH1_Haikei
                Exit Function
            End If
            If H1lb_SmpChr(wHaikei) = "" And wHaikei <> 0 Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(1) = "    ======== " & gInt & "çsñ⁄ îwåiêF =========   "
                     .MB_MSG(3) = "    ê›íËÇ≥ÇÍÇƒÇ¢ÇÈêFáÇÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢         "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                .Row = gInt
                Input_Area.Tag = gInt
                Call GridToText(Input_Area.Tag)
                Set ActObj = iH1_Haikei
                Exit Function
            End If
            
            ' ògêFáÇ
            If wWaku > 15 And wWaku <> 0 Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(1) = "    ========= " & gInt & "çsñ⁄ ògêF ==========   "
                     .MB_MSG(3) = "    êFáÇÇÕ ÇPÅ`ÇPÇTÇÃä‘ÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢         "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                .Row = gInt
                Input_Area.Tag = gInt
                Call GridToText(Input_Area.Tag)
                Set ActObj = iH1_Waku
                Exit Function
            End If
            If H1lb_SmpChr(wWaku) = "" And wWaku <> 0 Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(1) = "    ========= " & gInt & "çsñ⁄ ògêF ==========   "
                     .MB_MSG(3) = "    ê›íËÇ≥ÇÍÇƒÇ¢ÇÈêFáÇÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢         "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                .Row = gInt
                Input_Area.Tag = gInt
                Call GridToText(Input_Area.Tag)
                Set ActObj = iH1_Waku
                Exit Function
            End If
            
            If RTrim(.TextMatrix(gInt, 1)) <> "" And RTrim(.TextMatrix(gInt, 2)) = "" And RTrim(.TextMatrix(gInt, 4)) = "" Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(1) = "    ==================== " & gInt & "çsñ⁄ =====================          "
                     .MB_MSG(3) = "    éÛì¸Ç™ì¸óÕÇ≥ÇÍÇƒÇ¢ÇÈèÍçáÅAîwåiêFñîÇÕògêFÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢         "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                .Row = gInt
                Call vsGrid1_Click
                iH1_Haikei.SetFocus: Set ActObj = iH1_Haikei
                Exit Function
            End If
        
            If RTrim(.TextMatrix(gInt, 1)) = "" And _
             (RTrim(.TextMatrix(gInt, 2)) <> "" Or RTrim(.TextMatrix(gInt, 4)) <> "") Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(1) = "    ==================== " & gInt & "çsñ⁄ ====================="
                     .MB_MSG(3) = "    ògêFñîÇÕîwåiêFÇ™ì¸óÕÇ≥ÇÍÇƒÇ¢ÇÈèÍçáÅAéÛì¸Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢         "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                .Row = gInt
                Call vsGrid1_Click
                iH1_Ukeir.SetFocus: Set ActObj = iH1_Ukeir
                Exit Function
            End If
        
        Next gInt
        
    End With
'
    vsGrid1.Row = mSaveRow: vsGrid1.Col = 1
    Head1Chk = True
    Exit Function
Head1_Err:
    Exit Function
    
End Function
'**************************************
'*                                    *
'*          ÉfÅ[É^ÉxÅ[ÉXçXêV           *
'*                                    *
'**************************************
'*-------------------------------------------------*
'*      Éf Å[ É^ çX êV  äm îF ÉÅ Éb ÉZ Å[ ÉW
'*-------------------------------------------------*
Private Function CheckMessage() As Boolean
    CheckMessage = False
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     êFê›íËèÓïÒÇçXêVÇµÇ‹Ç∑Ç©ÅH      "
        .MB_Title = "çXêVämîF"
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call vsGrid1_Click
            Exit Function
        End If
    End With
    
    CheckMessage = True
End Function
'+----------------------+
'+    í«â¡•çXêV èàóù     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Static UpdFlg               As Boolean
    Dim wPutFlg                 As Boolean
'
'   ( ämîFÉÅÉbÉZÅ[ÉW )
    If Not CheckMessage Then Exit Function
    
    On Error GoTo DBPut_Err
'
    With CisDB
         .SQL = "delete ñºèÃÉ}ÉXÉ^ "
         .SQL = .SQL & " Where ãÊï™ñºèÃ = 'êFê›íË'"
         Call .DBExec
         
         .SQL = "delete êFê›íËÉeÅ[ÉuÉã"
         Call .DBExec
    End With

' êFê›íËèÓïÒçXêV
    For Each gObj In H1lb_SmpClr
        
        If Trim(H1lb_SmpChr(gObj.Index)) <> "" Then
            Call ItemsClearSYMCreate
            With SYMCreate
                .ãÊï™ñºèÃ = "êFê›íË"
                .ãÊï™åÖêî = 2
                .íl = Format(gObj.Index, "00")
                .ílñºèÃ = "êF" & Format(gObj.Index, "00")
                .êîéö1 = H1lb_SmpClr(gObj.Index).BackColor
                .êßå‰ãÊï™ = "X"
                .çÏê¨é“ = gTanto
                If gObj.Index = 0 Then
                   .ílñºèÃ = "ïWèÄ(îwåiêF)"
                End If
                If gObj.Index = 16 Then
                   .ílñºèÃ = "ïWèÄ(ògêF)"
                End If
                If Not SYMInitCreate(False) Then
                End If
            
            End With
        End If
    Next gObj

    With vsGrid1
        For gCnt1 = 1 To .Rows - 1
            If Trim(.TextMatrix(gCnt1, 1)) <> "" Then
                gStr = Trim(.TextMatrix(gCnt1, 0))
                If gStr <> "" Then
                    With IST
                        .éÛì¸ = Trim(vsGrid1.TextMatrix(gCnt1, 1))
                        .îwåiêF = CisFun.Val2(vsGrid1.TextMatrix(gCnt1, 2))
                        .ògêF = CisFun.Val2(vsGrid1.TextMatrix(gCnt1, 4))
                        Call ISTInsert
                    End With
                End If
            End If
        Next gCnt1
    End With
    
    On Error GoTo 0
    
    DBPut = True
    Call GridInit
    Call GridSet
    
    With CisFun
        .MB_Lines = 2
        .MB_MSG(1) = "   çXêVèIóπ          "
        .MB_Title = "çXêVèIóπämîF"
        .MB_Button = OK
        Call .MBOX
    End With
    Call GridSet
    
    Exit Function
    
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'=================================================================================
'
'           êFê›íËï\é¶
'
'
'=================================================================================
Private Sub InitColorDisp()

    For Each gObj In H1lb_SmpChr
        gObj.BackColor = &HFFFFFF
    Next gObj
    For Each gObj In H1lb_SmpClr
        gObj.ForeColor = &HFFFFFF
        gObj.BackColor = &HFFFFFF
    Next gObj

    gSL_Select = "Select * from ñºèÃÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " Where ãÊï™ñºèÃ = 'êFê›íË'"
    gSL_Select = gSL_Select & "   And ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " Order By íl"
    
    If SYMRead(gSL_Select) Then
       Do Until Not SYM_RDSTS
          gInt = CisFun.Val2(SYM.íl)
          H1lb_SmpChr(gInt) = "SAMPLE"
          H1lb_SmpChr(gInt).ForeColor = SYM.êîéö1
          H1lb_SmpClr(gInt).BackColor = SYM.êîéö1
          
          Call SYMReadNext
       Loop
    Else
       H1lb_SmpChr(0) = "SAMPLE"
       H1lb_SmpChr(0).ForeColor = 14737632
       H1lb_SmpClr(0).BackColor = 14737632
       H1lb_SmpChr(16) = "SAMPLE"
       H1lb_SmpChr(16).ForeColor = 0
       H1lb_SmpClr(16).BackColor = 0
    End If
    
    Call SYMClose

End Sub
'+==========================================================+
'+      ÉOÉäÉbÉhèàóù                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 0, "áÇ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 1, "éÛì¸")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 2, "êFNo")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 3, "îwåi")
        Call .FixedSet(ç∂íÜ, ç∂íÜ, Ç†ÇË, 4, 4, "êFNo.")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 5, "ògêF")
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0,6,"Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,4,5")
'  ï“èWì‡óeÉZÉbÉg
'       Call .EditSet(√∑ΩƒµÃﬁºﬁ™∏ƒ,∂◊—NO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       ÉOÉäÉbÉhì‡óeÉZÉbÉg        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
' íäèoèàóù èâä˙æØƒ

    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT * "
    gSL_Select = gSL_Select & "   FROM êFê›íËÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & "   Order by éÛì¸"
    
    If Not ISTRead(gSL_Select) Then
        Input_Area.Tag = 1
        Exit Sub
    End If
    
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With vsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
        ID = 0
        Do Until Not IST_RDSTS
           ID = ID + 1
            .Row = ID
            'áÇ
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            'éÛì¸
            .TextMatrix(ID, GridCol(1)) = IST.éÛì¸
            'îwåiêF
            If CisFun.Val2(IST.îwåiêF) > 9 Then
               gStr = ""
            Else
               gStr = " "
            End If
            .TextMatrix(ID, GridCol(2)) = gStr & Format(IST.îwåiêF, "##")
            Call GridColorSet(True, CisFun.Val2(IST.îwåiêF), 3)
            'ògêF
            If CisFun.Val2(IST.ògêF) > 9 Then
               gStr = ""
            Else
               gStr = " "
            End If
            .TextMatrix(ID, GridCol(4)) = gStr & Format(IST.ògêF, "##")
            Call GridColorSet(True, CisFun.Val2(IST.ògêF), 5)
            
            Call ISTReadNext
        Loop
        Call ISTClose
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
           
        'Call CisVsGrid.EditVisible
        wRow = .Row
        Call GridToText(vsGrid1.Row)
        .Redraw = True
    End With
    
' ∏ﬁÿØƒﬁì‡óeï\é¶å„ÅA√∑Ωƒì¸óÕ

End Sub
Private Sub GridToText(tRow As Integer)
    With vsGrid1
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, .CellTop, Input_Area.Width, .CellHeight
       'éÛì¸
       iH1_Ukeir = RTrim(.TextMatrix(tRow, 1))
       'îwåiêF
       iH1_Haikei = CisFun.Val2(.TextMatrix(tRow, 2))
       If CisFun.Val2(iH1_Haikei) > 0 And CisFun.Val2(iH1_Haikei) < 16 Then
          B1lb_Haikei.BackColor = H1lb_SmpClr(CisFun.Val2(iH1_Haikei)).BackColor
       Else
          B1lb_Haikei.BackColor = H1lb_SmpClr(0).BackColor
       End If
       'ògêF
       iH1_Waku = CisFun.Val2(.TextMatrix(tRow, 4))
       If CisFun.Val2(iH1_Waku) > 0 And CisFun.Val2(iH1_Waku) < 16 Then
          B1lb_Waku.BackColor = H1lb_SmpClr(CisFun.Val2(iH1_Waku)).BackColor
       Else
          B1lb_Waku.BackColor = H1lb_SmpClr(mColorMax + 1).BackColor
       End If
    End With
End Sub
Private Sub TextToGrid(tRow As Integer)
    With vsGrid1
        If RTrim(iH1_Ukeir) <> "" Or RTrim(iH1_Haikei) <> "" Or RTrim(iH1_Waku) <> "" Then
            'éÛì¸
            .TextMatrix(tRow, 1) = RTrim(iH1_Ukeir)
            'îwåiêF
            If CisFun.Val2(iH1_Haikei) > 9 Then
                gStr = ""
            Else
                gStr = " "
            End If
            .TextMatrix(tRow, 2) = gStr & Format(iH1_Haikei, "##")
            Call GridColorSet(True, CisFun.Val2(iH1_Haikei), 3, 1)
            'ògêF
            If CisFun.Val2(iH1_Waku) > 9 Then
                gStr = ""
            Else
                gStr = " "
            End If
            .TextMatrix(tRow, 4) = gStr & Format(iH1_Waku, "##")
            gLong = CisFun.Val2(iH1_Waku)
            If gLong = 0 Then .TextMatrix(tRow, 4) = "": gLong = 16
            Call GridColorSet(True, gLong, 5, 1)
        Else
            'îwåiêF
            .TextMatrix(tRow, 2) = ""
            Call GridColorSet(True, CisFun.Val2(0), 3, 1)
            'ògêF
            .TextMatrix(tRow, 4) = ""
            Call GridColorSet(True, CisFun.Val2(16), 5, 1)
        End If
    End With
End Sub

Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'ÉeÉLÉXÉgÇ©ÇÁÉOÉäÉbÉhÇ…ëﬁî
    Call TextToGrid(Val(Input_Area.Tag))
    gInt = vsGrid1.Row
    'Å™
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    'Å´
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid1.Rows - 1 Then
            If RTrim(vsGrid1.TextMatrix(gInt, 1)) <> "" Then
                vsGrid1.Rows = gInt + 1 + 1
                vsGrid1.TextMatrix(gInt + 1, 0) = gInt + 1
            Else
                GoTo Grid_RowChange_End
            End If
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid.DispMax
        End If
    End If
    'PageDown
    If KeyCode = vbKeyPageDown Then
        If gInt + CisVsGrid.DispMax > vsGrid1.Rows - 1 Then
            gInt = vsGrid1.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    vsGrid1.Row = gInt
'ÉOÉäÉbÉhÇ©ÇÁÉeÉLÉXÉgÇ…ï\é¶
    Call GridToText(vsGrid1.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Area.Visible = EnabledFlg
End Sub
'+----------------------+
'+  VSGrid1_Click
'+----------------------+
Private Sub vsGrid1_Click()
    Call TextToGrid(Val(Input_Area.Tag))
    Call GridTextEnabled(True)
    Call GridToText(vsGrid1.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
'+-----------------------+
'+  VSGrid_KeyDown
'+-----------------------+
Private Sub vsGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToGrid(Val(Input_Area.Tag))
    Call GridToText(vsGrid1.Row)
End Sub

