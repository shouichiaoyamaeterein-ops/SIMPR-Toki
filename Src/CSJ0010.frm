VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CSJ0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "取引先マスタメンテナンス"
   ClientHeight    =   11190
   ClientLeft      =   2925
   ClientTop       =   930
   ClientWidth     =   15240
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15240
   WindowState     =   2  '最大化
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Left            =   0
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox PNL_NewTehai 
      Height          =   870
      Left            =   8520
      ScaleHeight     =   810
      ScaleWidth      =   990
      TabIndex        =   20
      Top             =   3150
      Width           =   1050
      Begin Cis3D_v60.CIS3D Back_NewTehai 
         Height          =   810
         Left            =   0
         Top             =   0
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "手配№"
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
         Begin Cis3D_v60.CIS3D B1lb_Tehai 
            Height          =   360
            Left            =   90
            Top             =   390
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "999999"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -20
         End
      End
   End
   Begin VB.PictureBox B2_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   3900
      Left            =   2460
      ScaleHeight     =   3840
      ScaleWidth      =   10260
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4440
      Width           =   10320
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   3705
         Left            =   60
         Top             =   60
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   6535
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin vsFlexLib.vsFlexArray vsGrid 
            Height          =   3705
            Left            =   0
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   15
            Width           =   10095
            _Version        =   196608
            _ExtentX        =   17806
            _ExtentY        =   6535
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "No|機械№|機械名      |品番                      |工順|加工数|開始|終了|不良数|段取||"
            Rows            =   11
            Cols            =   25
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            BackColorAlternate=   12648384
            Begin Cis3D_v60.CIS3D Input_Area 
               Height          =   360
               Left            =   330
               Top             =   300
               Width           =   9570
               _ExtentX        =   16880
               _ExtentY        =   635
               BackColor       =   0
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
               Begin CisText_V60.CisText iB2_Hinbn 
                  Height          =   345
                  Left            =   2325
                  TabIndex        =   7
                  Top             =   0
                  Width           =   3255
                  _ExtentX        =   5741
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin CisText_V60.CisText iB2_Kikai 
                  Height          =   345
                  Left            =   0
                  TabIndex        =   6
                  Top             =   0
                  Width           =   825
                  _ExtentX        =   1455
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
               Begin CisText_V60.CisText iB2_Kojun 
                  Height          =   345
                  Left            =   5520
                  TabIndex        =   8
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   609
                  cFormat         =   "##"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
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
                  cFaZero         =   0
                  cFbComma        =   0
                  cGFormat        =   "##"
                  cILength        =   2
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB2_Kakou 
                  Height          =   345
                  Left            =   6120
                  TabIndex        =   9
                  Top             =   0
                  Width           =   825
                  _ExtentX        =   1455
                  _ExtentY        =   609
                  cFormat         =   "######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "######"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB2_Furyo 
                  Height          =   345
                  Left            =   8070
                  TabIndex        =   12
                  Top             =   0
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   609
                  cFormat         =   "######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "######"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB2_Syryo 
                  Height          =   345
                  Left            =   7500
                  TabIndex        =   11
                  Top             =   0
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
                  Text            =   "XXXX"
                  MaxLength       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
               Begin CisText_V60.CisText iB2_Kaisi 
                  Height          =   345
                  Left            =   6930
                  TabIndex        =   10
                  Top             =   0
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
                  Text            =   "XXXX"
                  MaxLength       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
               Begin CisText_V60.CisText iB2_Dandr 
                  Height          =   345
                  Left            =   8865
                  TabIndex        =   13
                  Top             =   0
                  Width           =   600
                  _ExtentX        =   1058
                  _ExtentY        =   609
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
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin Cis3D_v60.CIS3D B2lb_Kikai 
                  Height          =   330
                  Left            =   810
                  Top             =   0
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
         End
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   360
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   240
            _ExtentX        =   423
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
            Text            =   "X"
            MaxLength       =   1
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1860
      Left            =   2460
      ScaleHeight     =   1800
      ScaleWidth      =   5910
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2160
      Width           =   5970
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1680
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   5790
         _ExtentX        =   10213
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
         Begin Cis3D_v60.CIS3D Back_Torcd 
            Height          =   840
            Left            =   2100
            Top             =   0
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
            Begin CisText_V60.CisText iB1_Torcd 
               Height          =   360
               Left            =   90
               TabIndex        =   3
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
            Begin Cis3D_v60.CIS3D B1lb_TorNm 
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
            Left            =   0
            Top             =   840
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
            Begin CisText_V60.CisText iB1_Sagyo 
               Height          =   360
               Left            =   90
               TabIndex        =   4
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
            Begin Cis3D_v60.CIS3D B1lb_Sagyo 
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   840
            Left            =   0
            Top             =   0
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_SgYmd 
               Height          =   345
               Left            =   90
               TabIndex        =   2
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
         Begin Cis3D_v60.CIS3D UC_3D4 
            Height          =   840
            Left            =   3450
            Top             =   840
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   1482
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   840
            Left            =   2700
            Top             =   840
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "受入"
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
            Begin CisText_V60.CisText iB1_Ukeir 
               Height          =   360
               Left            =   180
               TabIndex        =   5
               Top             =   390
               Width           =   360
               _ExtentX        =   635
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
               Text            =   "XX"
               MaxLength       =   2
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   2460
      ScaleHeight     =   990
      ScaleWidth      =   4905
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   720
      Width           =   4965
      Begin Cis3D_v60.CIS3D Back_Tehai 
         Height          =   870
         Left            =   60
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1535
         BackColor       =   16761087
         ForeColor       =   16711680
         Caption         =   "手配№"
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
         Begin CisText_V60.CisText iH1_Tehai 
            Height          =   375
            Left            =   150
            TabIndex        =   0
            Top             =   390
            Width           =   990
            _ExtentX        =   1746
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
            cDataReplace    =   1
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hinbn 
         Height          =   870
         Left            =   1335
         Top             =   60
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "品　　番"
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
         Begin Cis3D_v60.CIS3D LABEL1 
            Height          =   300
            Left            =   525
            Top             =   60
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "作業日報管理行"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   390
            Width           =   3180
            _ExtentX        =   5609
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
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
      Height          =   1410
      Left            =   9150
      ScaleHeight     =   1350
      ScaleWidth      =   3570
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   3630
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   960
         Top             =   690
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
         Left            =   1800
         Top             =   690
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
         Left            =   2670
         Top             =   690
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   2670
         Top             =   60
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
         BtnType         =   4
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
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   255
      End
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   930
         Top             =   60
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
         BtnType         =   7
         BtnCaption      =   "新規"
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
         PFCaption       =   "( F1 )"
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
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   1800
         Top             =   60
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
         BtnType         =   8
         BtnCaption      =   "修正"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16776960
         BtnShadowColor  =   0
         PFCaption       =   "( F2 )"
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
         PFBackColor     =   16776960
      End
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   60
         Top             =   60
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
   End
   Begin VB.Timer Timer2 
      Left            =   720
      Top             =   420
   End
   Begin VB.Timer Timer1 
      Left            =   300
      Top             =   420
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
      Caption         =   "【 作業日報入力 】"
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'なし
         Height          =   345
         Index           =   1
         Left            =   10620
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   30
         Width           =   3675
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   345
            Left            =   2460
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8421631
            ForeColor       =   0
            Caption         =   "削除"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Delsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8421631
               ForeColor       =   0
               Caption         =   "ZZ9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   8421504
               cBoderColor2    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D6 
            Height          =   345
            Left            =   0
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8454143
            ForeColor       =   0
            Caption         =   "追加"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Addsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8454143
               ForeColor       =   0
               Caption         =   "ZZ9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   8421504
               cBoderColor2    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   345
            Left            =   1230
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   16777088
            ForeColor       =   0
            Caption         =   "修正"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Mntsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   16777088
               ForeColor       =   0
               Caption         =   "ZZ9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   8421504
               cBoderColor2    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
      End
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
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   8070
         Top             =   60
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
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
End
Attribute VB_Name = "CSJ0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   作業日報入力
'**       フォームID    :   CSKJ0010
'**       処理概要      :
'**
'**       作  成  日    :   2009/12/24  By CIS SIMPR-A
'**       変  更  日    :   2010/07/02  By CIS SIMPR-A 開始時間/終了時間 必須条件外し
'**       変  更  日    :   2010/07/07  By CIS SIMPR-A 品番重複チェック外し
'**       変  更  日    :   2011/07/27  By CIS SIMPR-A 手配桁数入力を6桁→7桁に変更
'**       変  更  日    :   2011/08/11  By CIS SIMPR-A 検索画面,手配桁数幅拡張
'**       変  更  日    :   2011/10/18  By CIS SIMPR-A 明細入力後のｾｯﾄﾌｫｰｶｽを担当者に移動
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim wChk            As Boolean

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer

    Dim mMaeHB          As String * 2   ' 直前の入力域
    Dim mTehaiNo        As String       ' 次に使用する作業日報管理№(手配№)
    Dim mRetRow         As Integer
    
    Const mDebugMode    As Boolean = False  ' True

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    mLostColor = iH1_Tehai.cLostColor
    
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 6      ' 未使用
        .InitGet = False
        .Init
    End With
    Call GridInit
    Input_Area.Visible = False
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    SyoriKB = "ADD"
    ProcHB = "B1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    iB1_SgYmd = Format(Now(), "yyyyMMdd")
    Call GetTehaiNo                         ' 起動直後の新手配NO

'-- 試験用 ---↓↓
If mDebugMode Then
Dim wWidthPlus As Integer: wWidthPlus = 10000
Me.Left = -120
Me.Width = Me.Width + 5000
B2_Area1.Left = 90
B2_Area1.Width = B2_Area1.Width + wWidthPlus
CIS3D6.Width = CIS3D6.Width + wWidthPlus
vsGrid.Width = vsGrid.Width + wWidthPlus
End If
'-- 試験用 ---↑↑
    
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
    Select Case Key_Code
           Case vbKeyF1:
                If PB_ADD.Visible Then
                    Call PB_ADD_Click    '【新規】
                End If
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub

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
   Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
   If ProcHB = "B2" Then
      If SyoriKB = "ADD" Then
         ProcHB = "B1"
      Else
         ProcHB = "H1"
      End If
     'ProcHB = "B1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B2")
      Call GridClear
      Call SyoriName(SyoriKB)
      PB_CAN.Visible = False
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
   If SyoriKB = "ADD" Then
      iB1_SgYmd.SetFocus
   Else
      iH1_Tehai.SetFocus
   End If
   Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    Dim wMae    As String
    wMae = ProcHB
    
    If SyoriKB <> "ADD" Then
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
        ProcHB = "B1"
        Call DispChange(ProcHB)
        If wMae <> "B1" Then
           Call iB1_SgYmd.SetFocus
        End If
        iB1_SgYmd = Format(Now(), "yyyyMMdd")
    End If
    
    PNL_NewTehai.Visible = True
    iH1_Hinbn.cLostColor = mLostColor
    B1lb_Tehai.Caption = mTehaiNo
 '   If wMae = "B1" Then Call ActObj.SetFocus
    iH1_Tehai = ""
    iH1_Hinbn = ""
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()

    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
    PNL_NewTehai.Visible = False
    iH1_Hinbn.Enabled = True
    iH1_Hinbn.cLostColor = mLostColor
   'iH1_Tehai.SetFocus
'    ActObj.SetFocus
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
        Call SyoriName(SyoriKB)
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
    PNL_NewTehai.Visible = False
    iH1_Hinbn.Enabled = False
    iH1_Hinbn.cLostColor = &HE0E0E0
   'iH1_Tehai.SetFocus
'    ActObj.SetFocus
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Tehai" Then Call Look_Tehai: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_SgYmd" Then
        If Len(Trim(iB1_SgYmd.Text)) < 4 Then iB1_SgYmd = ""
        iB1_SgYmd.ShowCalender
        Exit Sub
    End If
    If PB_Look.Tag = "iB1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iB1_Sagyo" Then Call Look_Sagyo: Exit Sub
    If PB_Look.Tag = "iB2_Kikai" Then Call Look_Kikai: Exit Sub
    If PB_Look.Tag = "iB2_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB2_Kakou" Then Call Look_Kakou: Exit Sub
    If PB_Look.Tag = "iB2_Kaisi" Then Call Look_Kaisi: Exit Sub
    If PB_Look.Tag = "iB2_Syryo" Then Call Look_Syryo: Exit Sub
    If PB_Look.Tag = "iB2_Kojun" Then Call Look_Kojun: Exit Sub     ' 2009.12.24 add (試験時のみ)
End Sub
'+----------------------------+
'+       手配№検索           +
'+----------------------------+
Private Sub Look_Tehai()
    RV_Left = 0
    RV_Top = 0
'    RV_TorKb = 8
'    RV_Call = "C"

    CKJ0010.Show vbModal
    Unload CKJ0010
    Set CKJ0010 = Nothing

    If RV_Rtn Then
        iH1_Tehai = RV_Tehai        ' 手配№
        iH1_Hinbn = RV_Hinbn        ' 品番
        iB1_SgYmd = Replace(RV_Ukeir, "/", "")  ' 年月日  ：受入で代用  ' 2010.01.15 upd : /が残っていたので修正した
        iB1_Torcd = RV_Torcd        ' 部門
        iB1_Sagyo = RV_Seban        ' 作業者  ：背番号で代用
        iB1_Ukeir = RVI_Ukeir       ' 受入                  ' 2010.01.06 add
        B1lb_TorNm = RVI_Torcd      ' 部門名
        B1lb_Sagyo = RVI_Seban      ' 作業者名：背番で代用
    End If
    RV_Tehai = ""                   ' 手配№
    RV_Hinbn = ""                   ' 品番
    RV_Ukeir = ""                   ' 年月日  ：受入で代用
    RV_Torcd = ""                   ' 部門
    RV_Seban = ""                   ' 作業者  ：背番号で代用
    RVI_Torcd = ""                  ' 部門名
    RVI_Seban = ""                  ' 作業者名：背番で代用
End Sub
'+----------------------------+
'+       品番検索             +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
   'RV_Call = "C"
    RVI_Hinmoku = 9
    RVI_SK = 9
    
   'CKK0030.Show vbModal        ' 品番検索時は0030
   'Unload CKK0030
   'Set CKK0030 = Nothing
    RVI_Torcd = Trim(iB1_Torcd)
    CKK0035.Show vbModal        ' 品番取引先検索時は0035
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       If ProcHB = "H1" Then
          iH1_Hinbn = RV_Hinbn
       Else
          iB2_Hinbn = RV_Hinbn
       End If
    End If
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
       iB1_Torcd = RV_Torcd
       Call iB1_Torcd_LostFocus
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
       iB1_Sagyo = RV_Code          ' 値
       B1lb_Sagyo = RV_Kubun        ' 値名称
    End If
End Sub
'+----------------------------+
'+       機械検索             +
'+----------------------------+
Private Sub Look_Kikai()
    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "機械NO"
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iB2_Kikai = RV_Code          ' 値
       B2lb_Kikai = RV_Kubun        ' 値名称
    End If
End Sub
'---------------------------------------------------' 2009.12.24 add start
'+----------------------------+
'+       工順検索             +
'+----------------------------+
Private Sub Look_Kojun()
    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "工順"
    CKJ0015.Show vbModal
    Unload CKJ0015
    Set CKJ0015 = Nothing

    If RV_Rtn Then
       iB2_Kojun = CStr(Rv_CallRtn)     ' 工順
       If Trim(iB2_Kikai) = "" Then                 ' 2009.12.28 add
          iB2_Kikai = RV_Code           ' 機械NO    ' 2009.12.28 add
          Call iB2_Kikai_LostFocus                  ' 2009.12.28 add
       End If                                       ' 2009.12.28 add
    End If
End Sub
'---------------------------------------------------' 2009.12.24 add end

'+----------------------------+
'+       加工数検索           +
'+----------------------------+
Private Sub Look_Kakou()
    MsgBox ("加工数検索")
End Sub
'+----------------------------+
'+       開始検索             +
'+----------------------------+
Private Sub Look_Kaisi()
    MsgBox ("開始検索")
End Sub
'+----------------------------+
'+       終了検索             +
'+----------------------------+
Private Sub Look_Syryo()
    MsgBox ("終了検索")
End Sub
'***********************************************************
'*                                                         *
'*      Focus 処理                                         *
'*                                                         *
'***********************************************************
'+---------------------------------------------------------+
'+          H1エリア
'+---------------------------------------------------------+
'+-----------------------------+
'+  手配№
'+-----------------------------+
Private Sub iH1_Tehai_GotFocus()
    Set ActObj = iH1_Tehai
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Tehai_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  品番(H1)
'+-----------------------------+
Private Sub iH1_Hinbn_GotFocus()
    Set ActObj = iH1_Hinbn
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+---------------------------------------------------------+
'+          B1エリア
'+---------------------------------------------------------+
'+-----------------------------+
'+  作業日
'+-----------------------------+
Private Sub iB1_SgYmd_GotFocus()
    Set ActObj = iB1_SgYmd
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SgYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  部門
'+-----------------------------+
Private Sub iB1_Torcd_GotFocus()
    Set ActObj = iB1_Torcd
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    B1lb_TorNm = ""
    If Trim(iB1_Torcd) <> "" Then
        gSL_Select = "SELECT * FROM 取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先CD = '" & Trim(iB1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 社内区分 = 1"
        If TRMRead(gSL_Select) Then
            B1lb_TorNm = TRM.正式名称
        Else
            B1lb_TorNm = ""
        End If
'-------------------------------------------------------------------------------------------------------' 2010.01.06 add start
        If Trim(iB1_Ukeir) = "" And ActiveControl.Name <> "iB1_Torcd" Then
            gSL_Select = "SELECT MIN(受入) AS 受入 FROM 取引先受入マスタ"
            gSL_Select = gSL_Select & " WHERE 取引先CD   = '" & Trim(iB1_Torcd) & "'"
            gSL_Select = gSL_Select & "   AND 取引先区分 = 1"
            If TUMRead(gSL_Select) Then
                iB1_Ukeir = TUM.受入
            End If
        End If
'-------------------------------------------------------------------------------------------------------' 2010.01.06 add end
    End If
End Sub
'+-----------------------------+
'+  作業者
'+-----------------------------+
Private Sub iB1_Sagyo_GotFocus()
    Set ActObj = iB1_Sagyo
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Sagyo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    B1lb_Sagyo = ""
    If Trim(iB1_Sagyo) <> "" Then
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '担当管理'"
        gSL_Select = gSL_Select & "   AND 値       = '" & Trim(iB1_Sagyo) & "'"
        If SYMRead(gSL_Select) Then
            B1lb_Sagyo = SYM.値名称
        Else
            B1lb_Sagyo = ""
        End If
    End If
End Sub
'+---------------------------------------------------------+
'+          グリッド編集エリア
'+---------------------------------------------------------+
'+-----------------------------+
'+  機械
'+-----------------------------+
Private Sub iB2_Kikai_GotFocus()
    PB_Look.Tag = ActiveControl.Name    ' PB_Look.Tag = "iB2_Kikai"
    PB_Look.Visible = True
    Set ActObj = iB2_Kikai
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB2_Kikai_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB2_Kikai_LostFocus()
   'PB_Look.Tag = ""
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    B2lb_Kikai = ""
    If Trim(iB2_Kikai) <> "" Then
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '機械NO'"
        gSL_Select = gSL_Select & "   AND 値       = '" & Trim(iB2_Kikai) & "'"
        If SYMRead(gSL_Select) Then
            B2lb_Kikai = SYM.値名称
        Else
            B2lb_Kikai = ""
        End If
    End If
End Sub
'+-----------------------------+
'+  品番(B2)
'+-----------------------------+
Private Sub iB2_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name    '     PB_Look.Tag = "iB2_Hinbn"
    PB_Look.Visible = True
    Set ActObj = iB2_Hinbn
End Sub
Private Sub iB2_Hinbn_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB2_Hinbn_LostFocus()
    
   'PB_Look.Tag = ""
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
'---------------------------------------------------------------------------------------' 2009.12.25 add start
    If Trim(iB2_Kojun) <> "" _
    Or Trim(iB2_Hinbn) = "" Then GoTo iB2_Hinbn_LostFocus_ED
        
    gSL_Select = "SELECT * FROM 品番工順マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & Trim(iB1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 機械NO = '" & Trim(iB2_Kikai) & "'"
    gSL_Select = gSL_Select & "   AND REPLACE(品番, '-', '') = '" & Trim(iB2_Hinbn) & "'"
   'gSL_Select = gSL_Select & "   AND 品番   = '" & Trim(iB2_Hinbn) & "'"
    If HKTRead(gSL_Select) Then
       iB2_Kojun = HKT.工順
       GoTo iB2_Hinbn_LostFocus_ED
    End If
        
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT MM1.値名称 機械NO名, MM2.値名称 機械名, HK.* FROM 品番工順マスタ HK"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ MM1"
    gSL_Select = gSL_Select & "   ON  MM1.区分名称 = '機械NO'"
    gSL_Select = gSL_Select & "   AND MM1.値       = HK.機械NO"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ MM2"
    gSL_Select = gSL_Select & "   ON  MM2.区分名称 = '機械'"
    gSL_Select = gSL_Select & "   AND MM2.値       = HK.機械"
    gSL_Select = gSL_Select & " WHERE HK.取引先 = '" & Trim(iB1_Torcd) & "'"
   'gSL_Select = gSL_Select & "   AND REPLACE(HK.品番, '-', '') = '" & Trim(iB2_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND HK.品番   = '" & Trim(iB2_Hinbn) & "'"
    If HKTRead(gSL_Select) Then
        If CisDB.RecordCount = 1 Then
           iB2_Kojun = HKT.工順
           If Trim(iB2_Kikai) = "" Then     ' 2009.12.28 add
              iB2_Kikai = HKT.機械NO        ' 2009.12.28 add
              Call iB2_Kikai_LostFocus      ' 2009.12.28 add
           End If                           ' 2009.12.28 add
           GoTo iB2_Hinbn_LostFocus_ED
        Else
           Call Look_Kojun
         ''iB2_Kojun = CStr(Rv_CallRtn)
           GoTo iB2_Hinbn_LostFocus_ED
        End If
    End If

iB2_Hinbn_LostFocus_ED:
'---------------------------------------------------------------------------------------' 2009.12.25 add start
End Sub
'---------------------------------------------------------------------------------------' 2009.12.24 del start
''+-----------------------------+
''+  工程
''+-----------------------------+
'Private Sub iB2_Kotei_GotFocus()
'    PB_Look.Tag = ""
'    PB_Look.Visible = False
'    Set ActObj = iB2_Kotei
'End Sub
'Private Sub iB2_Kotei_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call Grid_RowChange(KeyCode)
'End Sub
'Private Sub iB2_Kotei_LostFocus()
'    PB_Look.Tag = ""
'End Sub
'---------------------------------------------------------------------------------------' 2009.12.24 del end
'---------------------------------------------------------------------------------------' 2009.12.24 add start
'+-----------------------------+
'+  工順
'+-----------------------------+
Private Sub iB2_Kojun_GotFocus()
    PB_Look.Tag = ""        'ActiveControl.Name
    PB_Look.Visible = False 'True
    Set ActObj = iB2_Kojun
End Sub
Private Sub iB2_Kojun_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB2_Kojun_LostFocus()
    PB_Look.Tag = ""
End Sub
'---------------------------------------------------------------------------------------' 2009.12.24 add end
'+-----------------------------+
'+  加工数
'+-----------------------------+
Private Sub iB2_Kakou_GotFocus()
   'PB_Look.Tag = "iB2_Kakou"
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB2_Kakou
End Sub
Private Sub iB2_Kakou_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB2_Kakou_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  開始
'+-----------------------------+
Private Sub iB2_Kaisi_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB2_Kaisi
End Sub
Private Sub iB2_Kaisi_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB2_Kaisi_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  終了
'+-----------------------------+
Private Sub iB2_Syryo_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB2_Syryo
End Sub
Private Sub iB2_Syryo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB2_Syryo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  不良数
'+-----------------------------+
Private Sub iB2_Furyo_GotFocus()
    PB_Look.Tag = ""
    PB_Look.Visible = False
    Set ActObj = iB2_Furyo
End Sub
Private Sub iB2_Furyo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB2_Furyo_LostFocus()
    PB_Look.Tag = ""
End Sub
'+-----------------------------+
'+  段取
'+-----------------------------+
Private Sub iB2_Dandr_GotFocus()
    PB_Look.Tag = ""
    PB_Look.Visible = False
    Set ActObj = iB2_Dandr
    B_TabDummy.Tag = "E"
End Sub
Private Sub iB2_Dandr_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB2_Dandr_LostFocus()
    PB_Look.Tag = ""
End Sub
'+-----------------------------+
'+      B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If B_TabDummy.Tag = "S" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If B_TabDummy.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    SendKeys "{Tab}"
End Sub

'**************************************
'*                                    *
'*      処理区分により区分名表示      *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   
   SyoriKB = UCase(SYKB)                ' 2009.12.02 add
   Select Case SyoriKB
     Case "ADD"
          SyoriNM = "追加"
          K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
     Case "MNT"
          SyoriNM = "修正"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
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
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer2_Timer()
   If K_Sykbnm.Caption <> "" Then
      K_Sykbnm.Caption = ""
   Else
      K_Sykbnm.Caption = SyoriNM
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
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel
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
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
          If SyoriKB = "MNT" And Trim(iH1_Hinbn) = "" Then
             B1Mode = True
             B1Color = gPGotSel
          End If
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color
    
    PB_Look.Visible = True
    PB_ADD.Visible = Not B2Mode
    PB_MNT.Visible = Not B2Mode
    PB_DEL.Visible = Not B2Mode
    PB_CAN.Visible = B2Mode
    PB_END.Visible = Not B2Mode
    Select Case SyoriKB
        Case "ADD"
            PB_ADD.Visible = False
        Case "MNT"
            PB_MNT.Visible = False
        Case "DEL"
            PB_DEL.Visible = False
    End Select
    
    If H1Mode Then Call HeadBodyClear("B")
    If SyoriKB = "ADD" And B1Mode Then Call HeadBodyClear("H")
    
    Call GridTextEnabled(B2Mode)
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
'*--------------------------------------------------<< 2004/06/30 Insert Start >>
'           If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
'*--------------------------------------------------<< 2004/06/30 Insert End >>
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
'            If HeadBody = "B1" Then
'               If .Name Like "Back_*" Then .Visible = True
'            End If
        End With
    Next gObj
    If HeadBody = "B" Or HeadBody = "B1" Then
        CisVsGrid.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        Call GridSet
        vsGrid.Row = 1: Input_Area.Tag = "1"
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            Call GridInit
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            mMaeHB = ProcHB
            vsGrid.Row = 1: Input_Area.Tag = "1"
            Call VsGridToText(1)
           'ProcHB = "B1"
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call VsGridToText(vsGrid.Row)
        End If
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        mMaeHB = ProcHB
        ProcHB = "B2"
        Call DispChange(ProcHB)
        If SyoriKB <> "ADD" Then Call GridSet
        vsGrid.Row = 1: Input_Area.Tag = "1"
        Call VsGridToText(vsGrid.Row)
       'Call iB2_Kikai.SetFocus
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B2" Then
        mRetRow = CisFun.Val2(Input_Area.Tag)
        If Not Body2Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then
'---------------------------------------------------' 2009.12.03 del start
'            For Each gObj In B1Chk_Torkb
'                If gObj.Enabled Then
'                   gObj.SetFocus
'                   Exit For
'                End If
'            Next gObj
'---------------------------------------------------' 2009.12.03 del end
            GoTo ReturnPress_Ed
        End If
        
        If SyoriKB = "ADD" Then
           Call GetTehaiNo
        End If
        Call GridClear
        ProcHB = mMaeHB
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        vsGrid.Row = 1: Input_Area.Tag = "1"
                
        '================== 2011/10/18 Start
        If ProcHB = "B1" Then
            iB1_Sagyo.SetFocus
        End If
        '================== 2011/10/18 End
        
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
    
    If Trim(iH1_Tehai.Text) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    手配Noを入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Tehai.SetFocus
       Exit Function
    End If

    '-- データ確認 ---
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT * FROM 作業日報テーブル"
    gSL_Select = gSL_Select & " WHERE 作業日報管理NO = '" & Trim(iH1_Tehai) & "'"
    If Trim(iH1_Hinbn) <> "" Then
        gSL_Select = gSL_Select & "   AND 品番       = '" & Trim(iH1_Hinbn) & "'"
    End If
    '-- データ存在確認 -------
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT M1.値名称 AS 作業者名, M2.値名称 AS 機械NO名,"
    gSL_Select = gSL_Select & "       TM.正式名称 AS 部門名, "
    gSL_Select = gSL_Select & "       SN.*"
    gSL_Select = gSL_Select & " FROM  作業日報テーブル SN"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ M1"
    gSL_Select = gSL_Select & "            ON  M1.区分名称 = '担当管理'"
    gSL_Select = gSL_Select & "            AND M1.値       = RTRIM(SN.作業者)"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ M2"
    gSL_Select = gSL_Select & "            ON  M2.区分名称 = '機械NO'"
    gSL_Select = gSL_Select & "            AND M2.値       = RTRIM(SN.機械NO)"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM"
    gSL_Select = gSL_Select & "            ON  TM.取引先CD = RTRIM(SN.部門)"
    gSL_Select = gSL_Select & "            AND TM.社内区分 = 1"
    gSL_Select = gSL_Select & " WHERE 作業日報管理NO = '" & Trim(iH1_Tehai) & "'"
    If Trim(iH1_Hinbn) <> "" Then
        gSL_Select = gSL_Select & "   AND 品番       = '" & Trim(iH1_Hinbn) & "'"
    End If
    gSL_Select = gSL_Select & " ORDER BY SN.作業日報管理行"

    If Not SGTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当するデータはありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Tehai.SetFocus
       Exit Function
    End If
    iB1_SgYmd = SGT.作業日
    iB1_Torcd = SGT.部門
    iB1_Sagyo = SGT.作業者
    iB1_Ukeir = SGT.受入        ' 2010.01.06 add
    B1lb_TorNm = SGT.部門名
    B1lb_Sagyo = SGT.作業者名
    
    Head1Chk = True
End Function

'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    If Trim(iB1_SgYmd) = "" Or Len(Trim(iB1_SgYmd)) <> 8 _
    Or (Len(Trim(iB1_SgYmd)) = 8 And Mid(Trim(iB1_SgYmd), 7, 2) = "00") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    作業日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_SgYmd.SetFocus
       Exit Function
    End If
    If Trim(iB1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    部門を入力して下さい              "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_Torcd.SetFocus
       Exit Function
    End If
    If Trim(iB1_Sagyo) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    作業者を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_Sagyo.SetFocus
       Exit Function
    End If
    If Trim(iB1_Ukeir) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    受入を入力して下さい              "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_Ukeir.SetFocus
       Exit Function
    End If

    If SyoriKB <> "ADD" Then
        '-- データ存在確認 -------
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT M1.値名称 AS 作業者名, M2.値名称 AS 機械NO名,"
        gSL_Select = gSL_Select & "       TM.正式名称 AS 部門名, "
        gSL_Select = gSL_Select & "       SN.*"
        gSL_Select = gSL_Select & " FROM  作業日報テーブル SN"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ M1"
        gSL_Select = gSL_Select & "            ON  M1.区分名称 = '担当管理'"
        gSL_Select = gSL_Select & "            AND M1.値       = RTRIM(SN.作業者)"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ M2"
        gSL_Select = gSL_Select & "            ON  M2.区分名称 = '機械NO'"
        gSL_Select = gSL_Select & "            AND M2.値       = RTRIM(SN.機械NO)"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM"
        gSL_Select = gSL_Select & "            ON  TM.取引先CD = RTRIM(SN.部門)"
        gSL_Select = gSL_Select & "            AND TM.社内区分 = 1"
        gSL_Select = gSL_Select & " WHERE SN.作業日 = '" & Trim(iB1_SgYmd) & "'"
        gSL_Select = gSL_Select & "   AND SN.部門   = '" & Trim(iB1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND SN.作業者 = '" & Trim(iB1_Sagyo) & "'"
        gSL_Select = gSL_Select & "   AND SN.受入   = '" & Trim(iB1_Ukeir) & "'"        ' 2010.01.06 add
        gSL_Select = gSL_Select & "   AND 作業日報管理NO = '" & Trim(iH1_Tehai) & "'"
        If Trim(iH1_Hinbn) <> "" Then
            gSL_Select = gSL_Select & "   AND 品番       = '" & Trim(iH1_Hinbn) & "'"
        End If
        gSL_Select = gSL_Select & " ORDER BY SN.作業日報管理NO, SN.作業日報管理行"
        If Not SGTRead(gSL_Select) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    該当するデータがありません            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
    '        If ActObj.Visible Then ActObj.SetFocus
            Exit Function
        End If
    Else
        Call iB1_Torcd_LostFocus
        If B1lb_TorNm.Caption = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    部門　未登録            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            Call iB1_Torcd.SetFocus
            Exit Function
        End If
        Call iB1_Sagyo_LostFocus
        If B1lb_Sagyo.Caption = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    作業者　未登録            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            Call iB1_Sagyo.SetFocus
            Exit Function
        End If
'---------------------------------------------------------------------------------------------------' 2010.01.06 add start
        gSL_Select = "SELECT * FROM 取引先受入マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先CD   = '" & Trim(iB1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND 受入       = '" & Trim(iB1_Ukeir) & "'"
        If Not TUMRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    受入　未登録              "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Call iB1_Ukeir.SetFocus
        Exit Function
        End If
'---------------------------------------------------------------------------------------------------' 2010.01.06 add end
    End If

    Body1Chk = True
End Function

'****************************
'*      Body2ﾁｪｯｸ処理
'****************************
Private Function Body2Chk() As Boolean
    Dim wInput  As Boolean
    Dim wRowNo1 As Integer
    Dim wRowNo2 As Integer
    Dim wColNo1 As Integer
    Dim wRowCt1 As Integer
    Dim wRowCt2 As Integer
    Dim wSETime As String
    
    Body2Chk = False

    For wRowNo1 = 1 To vsGrid.Rows - 1          ' 計算領域クリア
        For wRowNo2 = 17 To vsGrid.Cols - 1
            vsGrid.TextMatrix(wRowNo1, wRowNo2) = ""
        Next
    Next
    
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call GridTextEnabled(False)

'ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'
    With vsGrid
'       << 行を詰める >>
        wInput = False
        For wRowNo1 = 1 To .Rows - 1
            '---------------------------------------------------------------------------------------' 2010.01.29 upd start
        ''''If RTrim(.TextMatrix(wRowNo1, 1)) = "" Then
            If RTrim(.TextMatrix(wRowNo1, 1)) = "" And RTrim(.TextMatrix(wRowNo1, 3)) = "" And _
               RTrim(.TextMatrix(wRowNo1, 4)) = "" And RTrim(.TextMatrix(wRowNo1, 5)) = "" And _
               RTrim(.TextMatrix(wRowNo1, 6)) = "" And RTrim(.TextMatrix(wRowNo1, 7)) = "" And _
               RTrim(.TextMatrix(wRowNo1, 8)) = "" And RTrim(.TextMatrix(wRowNo1, 9)) = "" _
            Then
            '---------------------------------------------------------------------------------------' 2010.01.29 upd end
                For wRowNo2 = wRowNo1 + 1 To .Rows - 1
                    '-------------------------------------------------------------------------------' 2010.01.29 upd start
                  ''If RTrim(.TextMatrix(wRowNo2, 1)) <> "" Then
                    If RTrim(.TextMatrix(wRowNo2, 1)) <> "" Or RTrim(.TextMatrix(wRowNo2, 3)) <> "" Or _
                       RTrim(.TextMatrix(wRowNo2, 4)) <> "" Or RTrim(.TextMatrix(wRowNo2, 5)) <> "" Or _
                       RTrim(.TextMatrix(wRowNo2, 6)) <> "" Or RTrim(.TextMatrix(wRowNo2, 7)) <> "" Or _
                       RTrim(.TextMatrix(wRowNo2, 8)) <> "" Or RTrim(.TextMatrix(wRowNo2, 9)) <> "" _
                    Then
                    '-------------------------------------------------------------------------------' 2010.01.29 upd end
                        For wColNo1 = 0 To .Cols - 1
                            .TextMatrix(wRowNo1, wColNo1) = .TextMatrix(wRowNo2, wColNo1)
                            .TextMatrix(wRowNo2, wColNo1) = ""
                        Next wColNo1
                        .TextMatrix(wRowNo1, 0) = Format(wRowNo1, "##")
'------------------------------------------------------------<< 2004/05/18 >> Insert Start
                        wInput = True
                        '> 行内容を移動時、入力域の再セット
                        If Val(Input_Area.Tag) = wRowNo1 Or _
                           Val(Input_Area.Tag) = wRowNo2 Then
                            Call VsGridToText(Val(Input_Area.Tag))
                        End If
'------------------------------------------------------------<< 2004/05/18 >> Insert End
                        Exit For
                    End If
                Next wRowNo2
            Else
                wInput = True
            End If
'            If Trim(.TextMatrix(wRowNo1, 1)) <> "" Then
'               .TextMatrix(wRowNo1, 11) = CStr(wRowNo1)
'            End If
        Next wRowNo1
        If Not wInput Then
            CisFun.MB_MSG(2) = "    明細を入力して下さい            "
            CisFun.MBOX
            Set ErrObj = iB2_Kikai
            wRowNo1 = 1
            GoTo Body2Chk_Err
        End If
'
        '   << 最終行NO取得 >>
        For wRowCt1 = 1 To .Rows - 1
            If wRowCt1 = .Rows - 1 Then Exit For
            If .TextMatrix(wRowCt1 + 1, 1) = "" Then Exit For
        Next
        
        For wRowNo1 = 1 To .Rows - 1
'============================= 2010/07/07 Start
            '---------------------------------------------------------------------------------------' 2010.01.29 upd start
          ''If RTrim(.TextMatrix(wRowNo1, 1)) = "" Then Exit For
            If RTrim(.TextMatrix(wRowNo1, 1)) = "" And RTrim(.TextMatrix(wRowNo1, 3)) = "" And _
               RTrim(.TextMatrix(wRowNo1, 4)) = "" And RTrim(.TextMatrix(wRowNo1, 5)) = "" And _
               RTrim(.TextMatrix(wRowNo1, 6)) = "" And RTrim(.TextMatrix(wRowNo1, 7)) = "" And _
               RTrim(.TextMatrix(wRowNo1, 8)) = "" And RTrim(.TextMatrix(wRowNo1, 9)) = "" _
            Then Exit For
            '---------------------------------------------------------------------------------------' 2010.01.29 upd end
            Input_Area.Tag = Str(wRowNo1)                       ' 2009.12.09 add
'        '   << 重複チェック >>
'            For wRowNo2 = wRowNo1 + 1 To .Rows - 1
'                If RTrim(.TextMatrix(wRowNo1, 1)) = RTrim(.TextMatrix(wRowNo2, 1)) And _
'                   RTrim(.TextMatrix(wRowNo1, 3)) = RTrim(.TextMatrix(wRowNo2, 3)) And _
'                   CisFun.Val2(.TextMatrix(wRowNo1, 4)) = CisFun.Val2(.TextMatrix(wRowNo2, 4)) Then
'                    CisFun.MB_MSG(1) = "    内容が重複しています。            "
'                    CisFun.MB_MSG(3) = "    " & wRowNo1 & "行目と " & wRowNo2 & "行目            "
'                    CisFun.MBOX
'                    Set ErrObj = iB2_Kikai
'                    wRowNo1 = wRowNo2
'                    GoTo Body2Chk_Err
'                End If
'            Next wRowNo2
'============================= 2010/07/07 End
        
        
        
        '   << 機械存在チェック >>
'======================================================= 2010/05/31 Start
'            If Trim(.TextMatrix(wRowNo1, 1)) = "" Then
'                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　機械　未入力！            "
'                CisFun.MBOX
'                Set ErrObj = iB2_Kikai
'                GoTo Body2Chk_Err
'            End If
'            gSL_Select = ""
'            gSL_Select = gSL_Select & "select *"
'            gSL_Select = gSL_Select & "  from 名称マスタ"
'            gSL_Select = gSL_Select & " where 区分名称 = '機械NO'"
'            gSL_Select = gSL_Select & "   and 値       = '" & Trim(.TextMatrix(wRowNo1, 1)) & "'"
'            If Not SYMRead(gSL_Select, 1) Then
'                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　機械　未登録！（名称マスタ）            "
'                CisFun.MBOX
'                Set ErrObj = iB2_Kikai
'                GoTo Body2Chk_Err
'            End If
            If Trim(.TextMatrix(wRowNo1, 1)) <> "" Then
                gSL_Select = ""
                gSL_Select = gSL_Select & "select *"
                gSL_Select = gSL_Select & "  from 名称マスタ"
                gSL_Select = gSL_Select & " where 区分名称 = '機械NO'"
                gSL_Select = gSL_Select & "   and 値       = '" & Trim(.TextMatrix(wRowNo1, 1)) & "'"
                If Not SYMRead(gSL_Select, 1) Then
                    CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　機械　未登録！（名称マスタ）            "
                    CisFun.MBOX
                    Set ErrObj = iB2_Kikai
                    GoTo Body2Chk_Err
                End If
            End If
'======================================================= 2010/05/31 End
        '   << 品番存在チェック >>
'
            If Trim(.TextMatrix(wRowNo1, 3)) = "" Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　品番　未入力！            "
                CisFun.MBOX
                Set ErrObj = iB2_Hinbn
                GoTo Body2Chk_Err
            End If
'            If Trim(.TextMatrix(wRowNo1, 1)) <> "" Or Trim(.TextMatrix(wRowNo1, 3)) <> "" Or Trim(.TextMatrix(wRowNo1, 4)) <> "" Or Trim(.TextMatrix(wRowNo1, 5)) <> "" Then
'               If Trim(.TextMatrix(wRowNo1, 3)) = "" Then
'                   CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　品番　未入力！            "
'                   CisFun.MBOX
'                   Set ErrObj = iB2_Hinbn
'                   GoTo Body2Chk_Err
'               End If
'            End If
            
            gSL_Select = ""
            gSL_Select = gSL_Select & "select *"
            gSL_Select = gSL_Select & "  from 品番取引先マスタ"
           'gSL_Select = gSL_Select & " where 品番   = '" & Trim(.TextMatrix(wRowNo1, 3)) & "'"     ' 2009.12.28 del
            gSL_Select = gSL_Select & " where REPLACE(品番, '-', '') = REPLACE('" & Trim(.TextMatrix(wRowNo1, 3)) & "', '-', '')"   ' 2009.12.28 add
           'gSL_Select = gSL_Select & "   and 取引先 = '" & Trim(iB1_Torcd) & "'"
            gSL_Select = gSL_Select & "   AND 取引先 = '" & Trim(iB1_Torcd) & "'"
            If Not HTMRead(gSL_Select, 1) Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　品番　未登録！（品番取引先マスタ）            "
                CisFun.MBOX
                Set ErrObj = iB2_Hinbn
                GoTo Body2Chk_Err
            End If
        '   << 品番工順存在チェック >>
            gSL_Select = ""
            gSL_Select = gSL_Select & "SELECT *"
            gSL_Select = gSL_Select & "  FROM 品番工順マスタ"
            gSL_Select = gSL_Select & " WHERE REPLACE(品番, '-', '') = REPLACE('" & Trim(.TextMatrix(wRowNo1, 3)) & "', '-', '')"   ' 2009.12.28 upd    ' 2009.12.25 upd
            gSL_Select = gSL_Select & "   AND 取引先 = '" & Trim(iB1_Torcd) & "'"
            If Trim(.TextMatrix(wRowNo1, 4)) <> "" Then                                             ' 2009.12.25 add
               gSL_Select = gSL_Select & "   AND 工順   =  " & Trim(.TextMatrix(wRowNo1, 4)) & " "  ' 2009.12.25 add
            End If                                                                                  ' 2009.12.25 add
          ''gSL_Select = gSL_Select & "   AND 機械NO = '" & Trim(.TextMatrix(wRowNo1, 1)) & "'"     ' 2009.12.25 del
            If Not HTMRead(gSL_Select, 1) Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　品番工順マスタ　未登録！           "
                CisFun.MBOX
                Set ErrObj = iB2_Hinbn
                GoTo Body2Chk_Err
            End If
'---------------------------------------------------------------------------------------------------' 2009.12.24 del start
'        '   << 工程チェック >>
'            If Trim(.TextMatrix(wRowNo1, 4)) = "" Then
'                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　工程を入力して下さい。            "
'                CisFun.MBOX
'                Set ErrObj = iB2_Kotei
'                GoTo Body2Chk_Err
'            End If
'---------------------------------------------------------------------------------------------------' 2009.12.24 del end
'---------------------------------------------------------------------------------------------------' 2009.12.24 add start
        '   << 工順チェック >>
            If Trim(.TextMatrix(wRowNo1, 4)) = "" Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　工順を入力して下さい。            "
                CisFun.MBOX
                Set ErrObj = iB2_Kojun
                GoTo Body2Chk_Err
            End If
'---------------------------------------------------------------------------------------------------' 2009.12.24 add end
    
        '   << 加工数チェック >>
            If CisFun.Val2(.TextMatrix(wRowNo1, 5)) = 0 Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　加工数を入力して下さい。            "
                CisFun.MBOX
                Set ErrObj = iB2_Kakou
                GoTo Body2Chk_Err
            End If
        
        
'========================== 2010/07/02 Del Start
'        '   << 開始チェック >>
'            If wRowNo1 = 1 And Trim(.TextMatrix(wRowNo1, 6)) = "" Then
'                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　開始を入力して下さい。            "
'                CisFun.MBOX
'                Set ErrObj = iB2_Kaisi
'                GoTo Body2Chk_Err
'            End If
'            If Trim(.TextMatrix(wRowNo1, 6)) <> "" Then
'                wSETime = Format(CisFun.Val2(.TextMatrix(wRowNo1, 6)), "0000")
'                If wSETime = "2400" Then wSETime = "0000"
'                .TextMatrix(wRowNo1, 6) = wSETime
'                If Not (("00" <= Mid(wSETime, 1, 2) And Mid(wSETime, 1, 2) <= "23" _
'                And "00" <= Mid(wSETime, 3, 2) And Mid(wSETime, 3, 2) <= "59")) Then
'                    CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　開始を正しく入力して下さい。            "
'                    CisFun.MBOX
'                    Set ErrObj = iB2_Kaisi
'                    GoTo Body2Chk_Err
'                End If
'            End If
'        '   << 終了チェック >>
'            If wRowNo1 = wRowCt1 And Trim(.TextMatrix(wRowNo1, 7)) = "" Then
'                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　終了を入力して下さい。            "
'                CisFun.MBOX
'                Set ErrObj = iB2_Syryo
'                GoTo Body2Chk_Err
'            End If
'            If Trim(.TextMatrix(wRowNo1, 7)) <> "" Then
'                wSETime = Format(CisFun.Val2(.TextMatrix(wRowNo1, 7)), "0000")
'                If wSETime = "2400" Then wSETime = "0000"
'                .TextMatrix(wRowNo1, 7) = wSETime
'                If Not (("00" <= Mid(wSETime, 1, 2) And Mid(wSETime, 1, 2) <= "23" _
'                And "00" <= Mid(wSETime, 3, 2) And Mid(wSETime, 3, 2) <= "59")) Then
'                    CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　終了を正しく入力して下さい。            "
'                    CisFun.MBOX
'                    Set ErrObj = iB2_Syryo
'                    GoTo Body2Chk_Err
'                End If
'            End If
'========================== 2010/07/02 Del End
        
        
        Next wRowNo1
    End With
'
'========================== 2010/07/02 Del Start
''   手配No毎の開始と終了を確認・算出する
'    If Not TimeSEChk(wRowNo1) Then
'       wRowNo1 = Input_Area.Tag
'       GoTo Body2Chk_Err
'    End If
'========================== 2010/07/02 Del End

    Body2Chk = True
    Exit Function
Body2Chk_Err:
    vsGrid.Row = wRowNo1
    Call VsGridToText(vsGrid.Row)
    Call GridTextEnabled(True)
    ErrObj.SetFocus
End Function

'-------------------------------------------'
'   手配No毎の開始と終了を確認・算出する    '
'-------------------------------------------'
Private Function TimeSEChk(ByRef wErr As Integer) As Boolean
    Dim wChkCtS As Integer
    Dim wChkCtE As Integer
    Dim wRowCt1 As Integer
    Dim wRowCt2 As Integer
    Dim wRowNo1 As Integer
    Dim wRowNo2 As Integer
    Dim wSETime As String
    Dim wWkCnt0 As Integer
    Dim wWkCnt1 As Integer
    Dim wWkCnt2 As Integer
    
    Dim wTehaiS As String
    Dim wTehaiI As Integer
    
    TimeSEChk = False
    
    With vsGrid
        wRowCt1 = 1: wRowCt2 = 1
        wRowNo1 = 1: wRowNo2 = 1
        wChkCtS = CisFun.Val2(.TextMatrix(1, 10))
         
        Do While 1 '''''''''wChkCtS <> 0
            For wRowNo2 = wRowNo1 To .Rows - 1
                If wRowNo2 >= .Rows - 1 Then wRowNo2 = .Rows - 1: Exit For
                If Trim(.TextMatrix(wRowNo2 + 1, 1)) = "" Then Exit For
                wChkCtE = CisFun.Val2(.TextMatrix(wRowNo2 + 1, 10))
                If SyoriKB = "MNT" And wChkCtS <> wChkCtE Then      'If SyoriKB = "MNT" And wChkCtS <> CisFun.Val2(.TextMatrix(wRowNo2 + 1, 10)) Then
                    Exit For
                End If
            Next wRowNo2
         
        '   << 開始～終了間に空白があるか確認する >>
            wWkCnt2 = 0
            For wWkCnt1 = wRowNo1 To wRowNo2
                If Trim(.TextMatrix(wWkCnt1, 6)) <> "" _
                Or Trim(.TextMatrix(wWkCnt1, 7)) <> "" Then
                    wWkCnt2 = 1
                    Exit For
                End If
            Next wWkCnt1
            
            ' wWkCnt2 がゼロでなければ開始or終了に空白があるので計算を行なう
            If wWkCnt2 <> 0 Then
                '-- 自動計算処理
                For wWkCnt1 = wRowNo1 To wRowNo2
                    .TextMatrix(wWkCnt1, 17) = ""   ' 基準工数
                    .TextMatrix(wWkCnt1, 18) = ""   ' 加工
                    .TextMatrix(wWkCnt1, 19) = ""   ' 割合
                    .TextMatrix(wWkCnt1, 20) = ""   ' 加工時間
                Next
                
                If Not GetSETime(wRowNo1, wRowNo2) Then
                    Call VsGridToText(CisFun.Val2(Input_Area.Tag))
                    Exit Function
                End If
            End If
            
            wRowNo1 = wRowNo2 + 1
            wChkCtS = wChkCtE
            If wRowNo1 > .Rows - 1 Then Exit Do
            If .TextMatrix(wRowNo1, 1) = "" Then Exit Do
        Loop
    End With
        
    TimeSEChk = True
End Function
'-------------------------------------------'
'   手配No毎の開始と終了を確認・算出する    '
'-------------------------------------------'
Private Function GetSETime(ByVal wGyoS As Integer, ByVal wGyoE As Integer) As Boolean
    Dim wCnt1 As Integer
    Dim wCnt2 As Integer
    Dim wCntS As Integer
    Dim wCntE As Integer
    Dim wStr1 As String
    Dim wTimS As String
    Dim wTime As String
    Dim wWkD1 As Double
    Dim wWkD2 As Double
    Dim wWkI1 As Integer
    Dim wWkI2 As Integer
    
    GetSETime = False

    If wGyoS <> wGyoE Then
        With vsGrid
            For wCnt1 = wGyoS To wGyoE
              If Trim(.TextMatrix(wCnt1, 6)) <> "" And Trim(.TextMatrix(wCnt1, 7)) <> "" Then
              Else
                Input_Area.Tag = Str(wCnt1)
                If Trim(.TextMatrix(wCnt1, 17)) = "" Then
                    If SyoriKB = "ADD" Then
                       wStr1 = Trim(iB1_Torcd)
                    Else
                       wStr1 = Trim(iB1_Torcd)
                    End If
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT * FROM 品番工順マスタ"
                    gSL_Select = gSL_Select & " WHERE REPLACE(品番,'-','') = REPLACE('" & Trim(.TextMatrix(wCnt1, 3)) & "', '-', '')"
                    gSL_Select = gSL_Select & "   AND 機械NO = '" & Trim(.TextMatrix(wCnt1, 1)) & "'"
                    gSL_Select = gSL_Select & "   AND 取引先 = '" & wStr1 & "'"
                    If Not HKTRead(gSL_Select) Then
                        CisFun.MB_MSG(2) = "    品番工順マスタ　未登録            "
                        CisFun.MBOX
                        Set ErrObj = iB2_Hinbn
                        Exit Function
                    End If
                    '-- 基準工数
                  ''.TextMatrix(wCnt1, 17) = Format(HKT.基準工数1 + HKT.基準工数2, "00.00")     ' 2010.01.22 del
                    '---------------------------------------------------------------------------' 2010.01.22 add start
                    If HKT.工数値選択 = 0 Or IsNull(HKT.工数値選択) = True _
                       Then .TextMatrix(wCnt1, 17) = Format(HKT.基準工数1, "00.00") _
                       Else .TextMatrix(wCnt1, 17) = Format(HKT.基準工数2, "00.00")
                    '---------------------------------------------------------------------------' 2010.01.22 add start
                    
                    If wCnt1 < wGyoE Then
                        For wCnt2 = wCnt1 To wGyoE
                            If .TextMatrix(wCnt1, 3) = .TextMatrix(wCnt2, 3) And _
                               .TextMatrix(wCnt1, 1) = .TextMatrix(wCnt2, 1) And _
                               .TextMatrix(wCnt1, 4) = .TextMatrix(wCnt2, 4) And _
                            (SyoriKB = "ADD" Or _
                            (SyoriKB = "MNT" And .TextMatrix(wCnt1, 13) = .TextMatrix(wCnt2, 13))) Then
                                .TextMatrix(wCnt2, 17) = .TextMatrix(wCnt1, 17)
                                '-- 加工数×基準工数
                                wWkD1 = CisFun.Val2(.TextMatrix(wCnt2, 17))
                                wWkI1 = CisFun.Val2(.TextMatrix(wCnt2, 5)) * wWkD1
                                .TextMatrix(wCnt2, 18) = CStr(wWkI1)
                            End If
                        Next wCnt2
                    End If
                End If
              End If
            Next wCnt1
        
            ' 時間算出
            wWkD2 = 0
            For wCnt1 = wGyoS To wGyoE
            wWkD2 = wWkD2 + CisFun.Val2(.TextMatrix(wCnt1, 18))
            Next
            If wWkD2 <> 0 Then
            For wCnt1 = wGyoS To wGyoE
            wWkD1 = CisFun.Val2(.TextMatrix(wCnt1, 18)) / wWkD2
            wWkI2 = wWkD1 * 100
            .TextMatrix(wCnt1, 19) = CStr(wWkI2)
            wWkI2 = CisFun.Val2(.TextMatrix(wCnt1, 18)) / 60
            .TextMatrix(wCnt1, 20) = CStr(wWkI2)
            Next
            End If
            
            Call SETime(wGyoS, wGyoE)
            
        End With
    End If

    GetSETime = True
End Function

'-------------------------------------------'
'   手配No毎の開始と終了を確認・算出する    '
'-------------------------------------------'
Private Function SETime(ByVal wGyoS As Integer, ByVal wGyoE As Integer) As Boolean
    Dim wCnt1 As Integer
    Dim wCnt2 As Integer
    Dim wCntS As Integer
    Dim wCntE As Integer
    Dim wStr1 As String
    Dim wTimS As String
    Dim wTime As String
    Dim wTimW As String
    
    SETime = False

wTimW = Mid(iB1_SgYmd, 5, 2) & "/" & Mid(iB1_SgYmd, 7, 2) & "/" & Mid(iB1_SgYmd, 1, 4) & " "
With vsGrid
    If wGyoS <> wGyoE Then
        wCntS = wGyoS
        wCntE = wGyoE
        
        For wCnt1 = wGyoS + 1 To wGyoE
            If Trim(.TextMatrix(wCnt1, 6)) = "" Then
                If Trim(.TextMatrix(wCnt1 - 1, 7)) <> "" Then
                    .TextMatrix(wCnt1, 6) = .TextMatrix(wCnt1 - 1, 7)
                End If
            End If
        Next wCnt1
        For wCnt1 = wGyoS To wGyoE - 1
            If Trim(.TextMatrix(wCnt1, 7)) = "" Then
                If Trim(.TextMatrix(wCnt1 + 1, 6)) <> "" Then
                    .TextMatrix(wCnt1, 7) = .TextMatrix(wCnt1 + 1, 6)
                End If
            End If
        Next

        wCntS = wGyoS
        wCntE = wGyoS
        Do While wCntE <= wGyoE
            
            For wCnt1 = wCntS To wGyoE
                If Trim(.TextMatrix(wCnt1, 6)) <> "" Then
                    wCntS = wCnt1
                    Exit For
                End If
            Next
            For wCnt2 = wCntE To wGyoE
                If Trim(.TextMatrix(wCnt2, 7)) <> "" Then
                    wCntE = wCnt2
                    Exit For
                End If
            Next
            If wCntS = wCntE Then
                wCntS = wCntS + 1
                wCntE = wCntE + 1
                GoTo SETime_LED
            End If
            
            For wCnt1 = wCntS To wCntE
                If wCnt1 = wCntS Then
                    wTimW = Mid(wTimW, 1, 11)
                    wTimW = wTimW & Mid(.TextMatrix(wCnt1, 6), 1, 2) & ":" & Mid(.TextMatrix(wCnt1, 6), 3, 4) & ":00"
                    .TextMatrix(wCnt1, 7) = Format(DateAdd("s", CisFun.Val2(.TextMatrix(wCnt1, 18)), wTimW), "hhmm")
                ElseIf wCnt1 = wCntE Then
                    .TextMatrix(wCnt1, 6) = .TextMatrix(wCnt1 - 1, 7)
                Else
                    .TextMatrix(wCnt1, 6) = .TextMatrix(wCnt1 - 1, 7)
                    wTimW = Mid(wTimW, 1, 11)
                    wTimW = wTimW & Mid(.TextMatrix(wCnt1, 6), 1, 2) & ":" & Mid(.TextMatrix(wCnt1, 6), 3, 4) & ":00"
                    .TextMatrix(wCnt1, 7) = Format(DateAdd("s", CisFun.Val2(.TextMatrix(wCnt1, 18)), wTimW), "hhmm")
                End If
            Next
            wCntS = wCnt1
            wCntE = wCnt1

SETime_LED:
        Loop

    End If
End With

End Function

'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function DBDelete()
    DBDelete = False

' ( 確認メッセージ )
'    K_Sykbnm.Caption = SyoriNM
'    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    削除処理を行います。            "
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Title = "削除確認"
        .MB_Button = OK_CAN
        If Not .MBOX Then
           Call HeadBodyClear("B1")
           Call GridInit
           Exit Function
        End If
    End With

    On Error GoTo DBDelete_Err
    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 削除処理 )
    Call DBDelete_SGT
'   件数表示
    iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")

    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0

    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True

' ( 確認メッセージ )
    With CisFun
        .MB_Lines = 5
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "    追加処理を行います。            "
            .MB_Title = "追加確認"
        Else
            .MB_MSG(2) = "    修正処理を行います。            "
            .MB_Title = "修正確認"
        End If
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            vsGrid.Row = mRetRow
            Input_Area.Tag = Str(mRetRow)
            Call VsGridToText(mRetRow)
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With

    On Error GoTo DBPut_Err
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    If SyoriKB <> "ADD" Then
    ' ( 削除処理 )
        If mMaeHB = "H1" Then Call DBDelete_SGT(, 0)    ' 2009.12.04 upd : If ～ Then を追加, 引数追加
    End If

' ( 追加処理 )
    If Not DBSet_SGT Then GoTo DBPut_Err
    For gIndex = 1 To vsGrid.Rows - 1
'        If RTrim(vsGrid.TextMatrix(gIndex, 1)) <> "" Then
        If RTrim(vsGrid.TextMatrix(gIndex, 3)) <> "" Then
            If mMaeHB <> "H1" And SyoriKB <> "ADD" Then ' 2009.12.04 add : B1からの修正時は
               Call DBDelete_SGT(, gIndex)              ' 2009.12.04 add : 該当するレコードを
            End If                                      ' 2009.12.04 add : 削除してから追加する
            Call DBInsert_SGT(, gIndex)
        End If
    Next gIndex
'   件数表示
    Select Case SyoriKB
        Case "ADD": iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
        Case "MNT": iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End Select

    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      削除処理 - 作業日報テーブル
'+------------------------------------+
Private Function DBDelete_SGT(Optional DBNo As Integer, Optional ID As Long = 0) As Boolean
    DBDelete_SGT = False
    
    If ID = 0 Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "DELETE FROM 作業日報テーブル"
        gSL_Select = gSL_Select & " WHERE 作業日報管理NO =  " & RTrim(iH1_Tehai) & " "
        If Trim(iH1_Hinbn) <> "" Then
           gSL_Select = gSL_Select & "   AND 品番        = '" & RTrim(iH1_Hinbn) & "'"
        End If
    Else
        gSL_Select = ""
        gSL_Select = gSL_Select & "DELETE FROM 作業日報テーブル"
        gSL_Select = gSL_Select & " WHERE 作業日報管理NO = " & CisFun.Val2(vsGrid.TextMatrix(ID, 10))
        gSL_Select = gSL_Select & "   AND 作業日報管理行 = " & CisFun.Val2(vsGrid.TextMatrix(ID, 11))
    End If

    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

    DBDelete_SGT = True
End Function
'+------------------------------------+
'+      セット - 作業日報テーブル
'+------------------------------------+
Private Function DBSet_SGT()
    DBSet_SGT = False

    DBSet_SGT = True
End Function
'+------------------------------------+
'+      追加処理 - 作業日報テーブル
'+------------------------------------+
Private Function DBInsert_SGT(Optional DBNo As Integer, Optional ID As Long) As Boolean
    DBInsert_SGT = False

    Call ItemsClearSGT
    With SGT
        If SyoriKB = "ADD" Then
            .作業日報管理NO = CisFun.Val2(B1lb_Tehai.Caption)
            .作業日報管理行 = ID
        Else
            .作業日報管理NO = CisFun.Val2(iH1_Tehai)                    ' .作業日報管理NO = CisFun.Val2(vsGrid.TextMatrix(ID, 10))
            .作業日報管理行 = CisFun.Val2(vsGrid.TextMatrix(ID, 11))    ' .作業日報管理行 = CisFun.Val2(vsGrid.TextMatrix(ID, 11))
        End If
        .作業日 = Trim(iB1_SgYmd)
        .部門 = Trim(iB1_Torcd)
        .受入 = Trim(iB1_Ukeir)                                         ' 2010.01.06 add
        .作業者 = Trim(iB1_Sagyo)
        .機械NO = Trim(vsGrid.TextMatrix(ID, 1))
        .品番 = Trim(vsGrid.TextMatrix(ID, 3))
        .工順 = CisFun.Val2(vsGrid.TextMatrix(ID, 4))                   ' 2009.12.25 upd : 工程→工順, Trim→CisFun.Val2
        .加工数 = CisFun.Val2(vsGrid.TextMatrix(ID, 5))
        .開始 = Trim(vsGrid.TextMatrix(ID, 6))
        .終了 = Trim(vsGrid.TextMatrix(ID, 7))
        If Trim(vsGrid.TextMatrix(ID, 8)) <> "" Then .不良数 = CisFun.Val2(vsGrid.TextMatrix(ID, 8))
        If Trim(vsGrid.TextMatrix(ID, 9)) <> "" Then .段取 = CisFun.Val2(vsGrid.TextMatrix(ID, 9))
    End With
    Call SGTInsert

    DBInsert_SGT = True
End Function

'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "No")
        Call .FixedSet(中中, 左中, あり, 6, 1, "機械No")        ' 2009.12.24 upd : 機械-->機械No / 桁数4-->8
        Call .FixedSet(中中, 左中, あり, 12, 2, "機械名")
        Call .FixedSet(中中, 左中, あり, 26, 3, "品　　番")
        Call .FixedSet(中中, 右中, あり, 4, 4, "工順")          ' 2009.12.25 upd : 工程→工順
        Call .FixedSet(中中, 右中, あり, 6, 5, "加工数")
        Call .FixedSet(中中, 右中, あり, 4, 6, "開始")
        Call .FixedSet(中中, 右中, あり, 4, 7, "終了")
        Call .FixedSet(中中, 右中, あり, 6, 8, "不良数")
        Call .FixedSet(中中, 右中, あり, 4, 9, "段取")
        '-- 記憶域
        Call .FixedSet(中中, 左中, あり, 6, 10, "手配No")
        Call .FixedSet(中中, 左中, あり, 6, 11, "行位置")
        Call .FixedSet(中中, 右中, あり, 8, 12, "作業日")
        Call .FixedSet(中中, 右中, あり, 7, 13, "部門")
        Call .FixedSet(中中, 右中, あり, 3, 14, "作業者")
        Call .FixedSet(中中, 右中, あり, 10, 15, "部門名")
        Call .FixedSet(中中, 右中, あり, 10, 16, "作業者名")
        Call .FixedSet(中中, 右中, あり, 7, 17, "基準工数")
        Call .FixedSet(中中, 右中, あり, 7, 18, "加工％")
        Call .FixedSet(中中, 右中, あり, 4, 19, "割合")
        Call .FixedSet(中中, 左中, あり, 8, 20, "加工時間")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0, 14, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9")
        If mDebugMode Then Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20"): vsGrid.ScrollBars = 3 ' 試験用
        .InitDisp
    End With
End Sub
'==========================================================='
'   グリッド移動
'==========================================================='
Private Sub Grid_RowChange(KeyCode As Integer)
    Dim wRow As String
    wRow = Input_Area.Tag
    
    If KeyCode <> vbKeyUp And _
       KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyPageUp And _
       KeyCode <> vbKeyPageDown Then Exit Sub

    If SyoriKB = "MNT" Then
       If KeyCode = vbKeyPageDown Then Exit Sub
       If KeyCode = vbKeyDown And _
         (Trim(vsGrid.TextMatrix(vsGrid.Row + 1, 0)) = "" _
       Or vsGrid.Row + 1 > vsGrid.Rows - 1) Then
          vsGrid.Row = CisFun.Val2(wRow)
          Input_Area.Tag = wRow
          Exit Sub
       End If
    End If

'   テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = vsGrid.Row
   '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
   '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid.Rows - 1 Then
            If RTrim(vsGrid.TextMatrix(gInt, 1)) <> "" Then
                vsGrid.Rows = gInt + 1 + 1
                vsGrid.TextMatrix(gInt + 1, 0) = gInt + 1
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
        If gInt + CisVsGrid.DispMax > vsGrid.Rows - 1 Then
            gInt = vsGrid.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    vsGrid.Row = gInt
'   グリッドからテキストに表示
    Call VsGridToText(vsGrid.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
'------------------------------------------------------------<< 2004/05/18 >> Update Start
'   Input_Area.Visible = EnabledFlg
'   B_TabDummy.Visible = EnabledFlg
'   *** ↑: 上記の１行は、入力処理時にグリッド移動が発生しないようにする為！ ***
'   *** カーソル位置が移動しない様にする為！ ***
    If EnabledFlg Then
        Input_Area.Visible = EnabledFlg
        B_TabDummy.Visible = EnabledFlg
    Else
        B_TabDummy.Visible = EnabledFlg
        Input_Area.Visible = EnabledFlg
    End If
'------------------------------------------------------------<< 2004/05/18 >> Update End
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid
        .TextMatrix(tRow, 1) = iB2_Kikai
'---------------------------------------------------------------------------------------' 2010.01.29 del start
'        If Trim(iB2_Kikai) = "" Then
'            For gInt = 2 To .Cols - 1
'                .TextMatrix(tRow, gInt) = ""
'            Next gInt
'            Exit Sub
'        End If
'---------------------------------------------------------------------------------------' 2010.01.29 del start
        
        .TextMatrix(tRow, 0) = tRow
        .TextMatrix(tRow, 1) = iB2_Kikai
        .TextMatrix(tRow, 2) = B2lb_Kikai
        .TextMatrix(tRow, 3) = iB2_Hinbn
        .TextMatrix(tRow, 4) = iB2_Kojun        ' 2009.12.24 upd : .TextMatrix(tRow, 4) = iB2_Kotei
        .TextMatrix(tRow, 5) = iB2_Kakou
        .TextMatrix(tRow, 6) = iB2_Kaisi
        .TextMatrix(tRow, 7) = iB2_Syryo
        .TextMatrix(tRow, 8) = iB2_Furyo
        .TextMatrix(tRow, 9) = iB2_Dandr
        
        If SyoriKB = "ADD" Then
           .TextMatrix(tRow, 10) = B1lb_Tehai
        Else
           .TextMatrix(tRow, 10) = iH1_Tehai
        End If
       '.TextMatrix(tRow, 11) = CisFun.Val2(Input_Area.Tag)
        .TextMatrix(tRow, 12) = iB1_SgYmd
        .TextMatrix(tRow, 13) = iB1_Torcd
        .TextMatrix(tRow, 14) = iB1_Sagyo
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        Input_Area.Tag = tRow                    ' 入力中の行位置を記憶
        Input_Area.Move Input_Area.Left, vsGrid.CellTop, Input_Area.Width, vsGrid.CellHeight
        B2lb_Kikai = Trim(.TextMatrix(tRow, 2))
        iB2_Kikai = Trim(.TextMatrix(tRow, 1))
        iB2_Hinbn = Trim(.TextMatrix(tRow, 3))
      ''iB2_Kotei = Trim(.TextMatrix(tRow, 4))  ' 2009.12.24 del
        iB2_Kojun = Trim(.TextMatrix(tRow, 4))  ' 2009.12.24 add
        iB2_Kakou = Trim(.TextMatrix(tRow, 5))
        iB2_Kaisi = Trim(.TextMatrix(tRow, 6))
        iB2_Syryo = Trim(.TextMatrix(tRow, 7))
        iB2_Furyo = Trim(.TextMatrix(tRow, 8))
        iB2_Dandr = Trim(.TextMatrix(tRow, 9))
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not SGT_RDSTS
        With vsGrid
            ID = ID + 1
            ' №
            .TextMatrix(ID, GridCol(0)) = Format(ID, "##")
            ' 機械
            .TextMatrix(ID, GridCol(1)) = SGT.機械NO
            ' 機械名
            .TextMatrix(ID, GridCol(2)) = SGT.機械NO名
            ' 品番
            .TextMatrix(ID, GridCol(3)) = SGT.品番
'            ' 工程                                         ' 2009.12.24 del
'            .TextMatrix(ID, GridCol(4)) = SGT.工程         ' 2009.12.24 del
            ' 工順                                          ' 2009.12.24 add
            .TextMatrix(ID, GridCol(4)) = SGT.工順          ' 2009.12.24 add
            ' 加工数
            .TextMatrix(ID, GridCol(5)) = Format(SGT.加工数, "######")
            ' 開始
            .TextMatrix(ID, GridCol(6)) = SGT.開始
            ' 終了
            .TextMatrix(ID, GridCol(7)) = SGT.終了
            ' 不良数
            .TextMatrix(ID, GridCol(8)) = Format(SGT.不良数, "######")
            ' 段取
            .TextMatrix(ID, GridCol(9)) = Format(SGT.段取, "###")
            
            ''''  記憶領域  ''''
            ' 手配No
            .TextMatrix(ID, GridCol(10)) = Format(SGT.作業日報管理NO, "#####0")
            ' 行位置
            .TextMatrix(ID, GridCol(11)) = Format(SGT.作業日報管理行, "#####0")
            ' 作業日
            .TextMatrix(ID, GridCol(12)) = SGT.作業日
            ' 部門
            .TextMatrix(ID, GridCol(13)) = SGT.部門
            ' 作業者
            .TextMatrix(ID, GridCol(14)) = SGT.作業者
            ' 部門名
            .TextMatrix(ID, GridCol(15)) = SGT.部門名
            ' 作業者名
            .TextMatrix(ID, GridCol(16)) = SGT.作業者名
        End With
        Call SGTReadNext
    Loop
    
    With vsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++     ﾃｷｽﾄ入力時      +++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And _
       KeyCode <> vbKeyDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(vsGrid.Row)
End Sub
'Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyUp And _
'       KeyCode <> vbKeyDown And _
'       KeyCode <> vbKeyPageUp And _
'       KeyCode <> vbKeyPageDown Then Exit Sub
'    Call TextToVsGrid(Val(Input_Area.Tag))
'    Call VsGridToText(vsGrid.Row)
'End Sub

Private Sub VSGrid_Click()
    Dim wRow As String
    Dim wCol As Integer
    wRow = Input_Area.Tag
    wCol = vsGrid.Col
        
    If vsGrid.Row = 0 Then
       vsGrid.Row = CisFun.Val2(wRow)
       Exit Sub
    End If
    
    If SyoriKB = "MNT" And vsGrid.TextMatrix(vsGrid.Row, 0) = "" Then
       vsGrid.Row = CisFun.Val2(wRow)
       Input_Area.Tag = wRow
       Call VsGridToText(vsGrid.Row)
       GoTo VsGrid_Click_Jmp
       Exit Sub
    End If
    
    If CisFun.Val2(Input_Area.Tag) = vsGrid.Row Then Exit Sub
    If Trim(vsGrid.TextMatrix(vsGrid.Row, 0)) = "" Then
       If (Trim(vsGrid.TextMatrix(vsGrid.Row - 1, 1)) = "") _
       Or (CisFun.Val2(Input_Area.Tag) = (vsGrid.Row - 1) _
       And Trim(iB2_Kikai) = "") Then
          Input_Area.Tag = wRow
          vsGrid.Row = CisFun.Val2(wRow)
          On Error Resume Next
          If ActObj.Visible Then ActObj.SetFocus
          On Error GoTo 0
          Exit Sub
       End If
    End If

    Call GridTextEnabled(True)
    Call TextToVsGrid(Val(Input_Area.Tag))
    If vsGrid.TextMatrix(vsGrid.Row, 0) = "" And _
    Trim(vsGrid.TextMatrix(vsGrid.Row - 1, 0)) = "" Then
       Input_Area.Tag = wRow
       On Error Resume Next
       If ActObj.Visible Then ActObj.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
'    iB2_Kikai = "": iB2_Hinbn = "": iB2_Kotei = "": iB2_Kakou = ""
'    iB2_Kaisi = "": iB2_Syryo = "": iB2_Furyo = "": iB2_Dandr = ""
'    B2lb_Kikai = ""
VsGrid_Click_Jmp:
    Call VsGridToText(vsGrid.Row)
    On Error Resume Next
    'If ActObj.Visible Then ActObj.SetFocus
    Select Case wCol
        Case 1
            iB2_Kikai.SetFocus
        Case 3
            iB2_Hinbn.SetFocus
        Case 4
            iB2_Kojun.SetFocus      ' 2009.12.25 upd : iB2_Kotei.SetFocus
        Case 5
            iB2_Kakou.SetFocus
        Case 6
            iB2_Kaisi.SetFocus
        Case 7
            iB2_Syryo.SetFocus
        Case 8
            iB2_Furyo.SetFocus
        Case 9
            iB2_Dandr.SetFocus
        Case Else
            iB2_Kikai.SetFocus
    End Select
    On Error GoTo 0
End Sub

'Private Sub VSGrid_Click()
''    Call GridTextEnabled(True)
'    Call TextToVsGrid(Val(Input_Area.Tag))
'    Call VsGridToText(vsGrid.Row)
'    On Error Resume Next
'    If ActObj.Visible Then ActObj.SetFocus
'    On Error GoTo 0
'End Sub

Private Sub GridClear()
    Dim Cnt1    As Integer
    Dim Cnt2    As Integer
    Dummy.Enabled = True: Dummy.SetFocus
    
    With vsGrid
        .Enabled = False
        .Rows = 11
        For Cnt1 = 1 To .Rows - 1
            For Cnt2 = 0 To .Cols - 1
                .TextMatrix(Cnt1, Cnt2) = ""
            Next Cnt2
        Next Cnt1
        .Enabled = True
    End With
    Call GridInit
    iB2_Kikai = "": iB2_Hinbn = "": iB2_Kojun = "": iB2_Kakou = "": B2lb_Kikai = ""     ' 2009.12.24 upd : iB2_Kotei = "": を削除、工順を追加
    
    iB2_Kaisi = "": iB2_Syryo = "": iB2_Furyo = "": iB2_Dandr = ""
    Input_Area.Tag = "1"
    Input_Area.Move Input_Area.Left, vsGrid.CellTop, Input_Area.Width, vsGrid.CellHeight

    Dummy.Enabled = False
End Sub

'-------------------------------------------'
'   手配Noの獲得                            '
'-------------------------------------------'
Private Function GetTehaiNo() As Boolean
    GetTehaiNo = False
    
    Call ItemsClearSGT
    With CisDB
        .SQL = "番号獲得"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = OutPut用
        .ParaNo = 3: .ParaIO = Return用
    '
        .ParaNo = 1: .ParaValue = 21
    '
        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    手配NOの獲得に失敗しました            "
            CisFun.MB_Title = "採番処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Function
        Else
            .ParaNo = 2: SGT.作業日報管理NO = .ParaValue
        End If
    End With
    B1lb_Tehai.Caption = ""
    If Trim(SGT.作業日報管理NO) <> "" Then
       B1lb_Tehai.Caption = SGT.作業日報管理NO
       mTehaiNo = SGT.作業日報管理NO
    End If

    GetTehaiNo = True
End Function

