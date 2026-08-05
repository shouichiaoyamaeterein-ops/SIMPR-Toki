VERSION 5.00
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form CXM2200 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "得意先スクラップ単価メンテ"
   ClientHeight    =   11190
   ClientLeft      =   1785
   ClientTop       =   420
   ClientWidth     =   15360
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
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.Timer Timer2 
      Left            =   3930
      Top             =   1350
   End
   Begin VB.Timer Timer1 
      Left            =   1950
      Top             =   1620
   End
   Begin VB.TextBox Dummy 
      Height          =   360
      Left            =   1170
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox B2_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   9195
      ScaleHeight     =   3480
      ScaleWidth      =   5955
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   6015
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   360
         Left            =   210
         Top             =   3870
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
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   465
         Left            =   90
         Top             =   90
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   820
         BackColor       =   12648384
         ForeColor       =   16711680
         Caption         =   "【スクラップ単価履歴照会】"
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
         cAlingnment     =   7
         cPositionX      =   60
      End
      Begin VB.PictureBox CIS3D30 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         ScaleHeight     =   360
         ScaleWidth      =   5730
         TabIndex        =   9
         Top             =   585
         Width           =   5790
         Begin Cis3D_v60.CIS3D B2lb_Kensu 
            Height          =   315
            Left            =   4650
            Top             =   15
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            BackColor       =   16761024
            Caption         =   "999件"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D B2lb_Zai 
            Height          =   315
            Left            =   15
            Top             =   15
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            BackColor       =   16777152
            Caption         =   "材質　：　鉄"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cPositionX      =   60
         End
      End
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CXM2200.frx":0000
         Height          =   2370
         Left            =   90
         OleObjectBlob   =   "CXM2200.frx":0015
         TabIndex        =   15
         Top             =   1050
         Width           =   5760
      End
   End
   Begin VB.PictureBox H1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4155
      ScaleHeight     =   1275
      ScaleWidth      =   6810
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2940
      Width           =   6870
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1185
         Left            =   60
         Top             =   45
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   2090
         ForeColor       =   16711680
         Caption         =   "対象期"
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
         Begin Cis3D_v60.CIS3D H1lb_Kikan 
            Height          =   285
            Left            =   120
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   503
            BackColor       =   16777152
            Caption         =   "YYYY/MM/DD～YYYY/MM/DD"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   345
            Left            =   750
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "年"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Y 
            Height          =   345
            Left            =   120
            TabIndex        =   0
            Top             =   450
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   609
            cFormat         =   "####"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "####"
            MaxLength       =   4
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "####"
            cILength        =   4
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1050
            TabIndex        =   10
            Top             =   330
            Width           =   1725
            Begin VB.OptionButton H1Op_Ki 
               BackColor       =   &H00C0C0C0&
               Caption         =   "下期"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   885
               TabIndex        =   2
               Top             =   165
               Width           =   795
            End
            Begin VB.OptionButton H1Op_Ki 
               BackColor       =   &H00C0C0C0&
               Caption         =   "上期"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   90
               TabIndex        =   1
               Top             =   165
               Value           =   -1  'True
               Width           =   795
            End
         End
      End
      Begin Cis3D_v60.CIS3D H1_TorMsg 
         Height          =   1185
         Left            =   3015
         Top             =   45
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   2090
         ForeColor       =   16711680
         Caption         =   "得  意  先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
            Height          =   375
            Left            =   135
            TabIndex        =   3
            Top             =   525
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
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1110
            Top             =   525
            Width           =   2475
            _ExtentX        =   4366
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
   Begin VB.PictureBox B1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   4155
      ScaleHeight     =   5070
      ScaleWidth      =   6810
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4290
      Width           =   6870
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   4980
         Left            =   75
         Top             =   60
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   8784
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
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   4050
            Left            =   165
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   825
            Width           =   6345
            _Version        =   196608
            _ExtentX        =   11192
            _ExtentY        =   7144
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "№ |材質               |前回単価     |今回単価    |材質|履歴NO"
            Rows            =   11
            Cols            =   7
            BackColor       =   12648384
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   360
            GridLines       =   2
            SelectionMode   =   1
            BackColorAlternate=   12648384
            Begin Cis3D_v60.CIS3D Input_Area 
               Height          =   375
               Left            =   4455
               Top             =   345
               Visible         =   0   'False
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   661
               BackColor       =   4210752
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
               Begin VB.TextBox F_TabDummy 
                  Height          =   330
                  Left            =   2505
                  TabIndex        =   12
                  Text            =   "Text1"
                  Top             =   0
                  Width           =   240
               End
               Begin VB.TextBox B_TabDummy 
                  Height          =   330
                  Left            =   -240
                  TabIndex        =   13
                  Text            =   "Text1"
                  Top             =   0
                  Width           =   240
               End
               Begin CisText_V60.CisText iB1_Tan 
                  Height          =   375
                  Left            =   0
                  TabIndex        =   14
                  Top             =   0
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   661
                  cFormat         =   "#,###,##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#,###,##0.00"
                  MaxLength       =   12
                  cDataType       =   2
                  cDataReplace    =   1
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "######0.00"
                  cILength        =   7
               End
            End
         End
         Begin CisBtn_60.CisBtn PB_RIRE 
            Height          =   585
            Left            =   5685
            Top             =   165
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1032
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PFSize          =   0
            BtnType         =   11
            BtnCaption      =   "履歴"
            BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BtnForeColor    =   16744703
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
            PFBackColor     =   16744703
         End
         Begin Cis3D_v60.CIS3D PB_Eikyo 
            Height          =   600
            Left            =   4425
            Top             =   150
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   1058
            ForeColor       =   16711680
            Caption         =   "影響額"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
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
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   9120
      ScaleHeight     =   750
      ScaleWidth      =   6030
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6090
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   30
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3450
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
         Left            =   4305
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
         Left            =   885
         Top             =   90
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
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   1740
         Top             =   90
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   2595
         Top             =   90
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
      Caption         =   "【得意先スクラップ単価 メンテ  】"
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
         Left            =   14220
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
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
         Height          =   285
         Left            =   11595
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
End
Attribute VB_Name = "CXM2200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   得意先スクラップ単価　メンテ
'**       フォームID    :   CXM2200
'**       処理概要      :
'**
'**       作  成  日    :   2008/05/15  By CIS
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    Dim mTanKb          As Byte
    Dim ActObj          As Object
    Dim SaveRow         As Integer
    
    Dim SY_KamiM        As String       ' 上期開始月
    Dim SY_KamiSMD      As String       ' 上期開始日付
    Dim SY_KamiEMD      As String       ' 上期終了日付
    Dim SY_SimoM        As String       ' 下期開始月
    Dim SY_SimoSMD      As String       ' 下期開始日付
    Dim SY_SimoEMD      As String       ' 下期終了日付
    
    Dim wKi             As String       ' 期
    
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub form_load()
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
    
    
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
'
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 360
        .PatanMax = 1
        .SelectCol = 6
        .InitGet = False
        .Init
    End With
    Call GridInit
    
'   #--------------------------#
'   #    名称マスタ初期登録    #
'   #--------------------------#
    Call MeisyoInit
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
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    H1Op_Ki(0).Value = True
    Adodc1.ConnectionString = CisDB.ConnectString
    
    FormAct = Not FormAct
'
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
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF6:       If ProcHB = "B1" Then Call PB_RIRE_Click    '【履歴】
           Case vbKeyF7:       If PB_LOOK.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
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
    '
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
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      Call Head1Chk             ' 2008.08.29 add
    iH1_Y.SetFocus
      iB1_Tan.Tag = ""
      GoTo PB_CAN_Ed
   End If
   
   If ProcHB = "B2" Then
      ProcHB = "B1"
      Call DispChange(ProcHB)
      vsGrid1.Row = 1
      iB1_Tan.Tag = ""
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
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
    iH1_Y.SetFocus
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    Call iH1_Y_LostFocus        ' 2008.08.29 add
    gSL_Select = " SELECT 対象年, 対象期, 得意先"
    gSL_Select = gSL_Select & " FROM 得意先スクラップ単価マスタ"
    gSL_Select = gSL_Select & " WHERE 対象年 + CAST(対象期 AS VARCHAR) + 得意先 <"
    gSL_Select = gSL_Select & " '" & iH1_Y & "' + '" & wKi & "' + '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " GROUP BY 対象年, 対象期, 得意先"
    gSL_Select = gSL_Select & " ORDER BY 対象年 + CAST(対象期 AS VARCHAR) + 得意先 Desc "
    If Not TSMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     先頭レコードです            "
            .MB_Title = "【後退】"
            .MB_Button = OK
            .MBOX
        End With
        Call CisFrm.MousePT(1)
        Exit Sub
    End If
        
    iH1_Y = TSM.対象年
    H1Op_Ki(TSM.対象期).Value = True
    Call DateDSP(CInt(TSM.対象期))
    
    iH1_Torcd = TSM.得意先
    Call iH1_Torcd_LostFocus
    
    Call DBInput("")
    
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Y.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    Call iH1_Y_LostFocus        ' 2008.08.29 add
    gSL_Select = " SELECT 対象年, 対象期, 得意先"
    gSL_Select = gSL_Select & " FROM 得意先スクラップ単価マスタ"
    gSL_Select = gSL_Select & " WHERE 対象年 + CAST(対象期 AS VARCHAR) + 得意先 >"
    gSL_Select = gSL_Select & " '" & iH1_Y & "' + '" & wKi & "' + '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " GROUP BY 対象年, 対象期, 得意先"
    gSL_Select = gSL_Select & " ORDER BY 対象年 + CAST(対象期 AS VARCHAR) + 得意先"

    If Not TSMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     最終レコードです            "
            .MB_Title = "【前進】"
            .MB_Button = OK
            .MBOX
        End With
        Call CisFrm.MousePT(1)
        Exit Sub
    End If
    
    iH1_Y = TSM.対象年
    H1Op_Ki(TSM.対象期).Value = True
    Call DateDSP(CInt(TSM.対象期))
    
    iH1_Torcd = TSM.得意先
    Call iH1_Torcd_LostFocus
        
    Call DBInput("")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Y.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_LOOK.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
End Sub
'+----------------------------+
'+    履  歴  キ ー(F7)       +
'+----------------------------+
Private Sub PB_RIRE_Click()
    ProcHB = "B2"
    Call DispChange(ProcHB)
    
    Call RirekiDisp
End Sub
'+----------------------+
'+     影響額ボタン     +
'+----------------------+
Private Sub PB_Eikyo_Click()
    RV_Torcd = RTrim(iH1_Torcd)
    RV_Zaisyu = RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 4))
    RV_Y = RTrim(iH1_Y)
    RV_Ki = CisFun.Val2(wKi)
    Rv_Call3 = "X"
    
    CXM2240.Show vbModal
    Unload CXM2240
    Set CXM2240 = Nothing
End Sub
'+----------------------------+
'+       得意先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 2
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+  iH1_Y
'+-----------------------------+
Private Sub iH1_Y_LostFocus()
    If Len(Trim(iH1_Y)) = 1 Then iH1_Y = "200" & Trim(iH1_Y)  ' 2008.08.29 add
    Call DateDSP(CisFun.Val2(wKi))
End Sub
'+-----------------------------+
'+  iH1_Torcd
'+-----------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    H1lb_Tornm = ""
    If TorNmGet(iH1_Torcd, 2, 1) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------+
'+     H1Op_Ki     +
'+-----------------+
Private Sub H1Op_Ki_Click(Index As Integer)
    wKi = Index
    
    Call DateDSP(Index)
    
End Sub
'+-----------------------------+
'+  iB1_Tan
'+-----------------------------+
Private Sub iB1_Tan_GotFocus()
    Set ActObj = iB1_Tan
    iB1_Tan.Tag = "S"
End Sub
Private Sub iB1_Tan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Tan_LostFocus()
    If CisFun.Val2(iB1_Tan) <> 0 Then
        vsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 3) = iB1_Tan
    Else
        vsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 3) = ""
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If iB1_Tan.Tag <> "" Then
       Call TextToVsGrid(Val(Input_Area.Tag))
       Call Grid_RowChange(vbKeyUp)
    End If
    iB1_Tan.SetFocus
End Sub

'+-----------------------------+
'+  F_TabDummy
'+-----------------------------+
Private Sub F_TabDummy_GotFocus()
    If iB1_Tan.Tag <> "" Then
       Call TextToVsGrid(Val(Input_Area.Tag))
       Call Grid_RowChange(vbKeyDown)
    End If
    iB1_Tan.SetFocus
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
     Case Else
          SyoriNM = "修正"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
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
   Call CisFrm.TimeSet(1)
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
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color
    B2_Area1.Visible = B2Mode
    PB_LOOK.Visible = False
    PB_DEL.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_ENT.Visible = Not B2Mode
    
    Call GridTextEnabled(B1Mode)
    
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
            If .Name Like wbc_Enm & "*Ch_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
       On Error Resume Next
       Call CisVsGrid.Clear
       On Error GoTo 0
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
            vsGrid1.Row = 1
            
            Call VsGridToText(vsGrid1.Row)
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        Call Head1Chk           ' 2008.08.29 add
        iB1_Tan.Tag = ""
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
    
    If RTrim(iH1_Y) = "" Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    対象年を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
       iH1_Y.SetFocus
       Exit Function
    End If
        
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    得意先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 2) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    得意先　未登録！            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    
' マスタ入力 SELECT文 ｾｯﾄ
    If DBInput("") Then
    Else
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "  該当するデータが存在しません。     "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Y.SetFocus
        Exit Function
    End If
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
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
    
    Call HeadBodyClear("B")
    
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT ISNULL(SY1.値名称, '') 材質名, ISNULL(SY1.値, '') 材質," & vbCrLf
    gSL_Select = gSL_Select & "        ISNULL(TSM.単価, 0) 単価, ISNULL(TSM.履歴NO, 0) 履歴NO" & vbCrLf
    gSL_Select = gSL_Select & "   FROM 名称マスタ SY1" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 得意先スクラップ単価マスタ TSM" & vbCrLf
    gSL_Select = gSL_Select & "             ON TSM.材質 = SY1.値" & vbCrLf
    gSL_Select = gSL_Select & "            AND TSM.得意先 = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
'    If H1Op_Ki(0).Value Then
'        gSL_Select = gSL_Select & "            AND TSM.適用開始日付 = '" & SY_KamiSMD & "'" & vbCrLf
'        gSL_Select = gSL_Select & "            AND TSM.適用終了日付 = '" & SY_KamiEMD & "'" & vbCrLf
'    Else
'        gSL_Select = gSL_Select & "            AND TSM.適用開始日付 = '" & SY_SimoSMD & "'" & vbCrLf
'        gSL_Select = gSL_Select & "            AND TSM.適用終了日付 = '" & SY_SimoEMD & "'" & vbCrLf
'    End If
    gSL_Select = gSL_Select & "            AND TSM.履歴NO = (SELECT MAX(履歴NO)" & vbCrLf
    gSL_Select = gSL_Select & "                                FROM 得意先スクラップ単価マスタ TSM1" & vbCrLf
    gSL_Select = gSL_Select & "                               WHERE TSM1.得意先 = TSM.得意先" & vbCrLf
'    gSL_Select = gSL_Select & "                                 AND TSM1.適用開始日付 = TSM.適用開始日付" & vbCrLf
'    gSL_Select = gSL_Select & "                                 AND TSM1.適用終了日付 = TSM.適用終了日付" & vbCrLf
    gSL_Select = gSL_Select & "                                 AND TSM1.材質 = TSM.材質)" & vbCrLf
    gSL_Select = gSL_Select & " WHERE SY1.区分名称 = '材種区分'" & vbCrLf
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''" & vbCrLf
'Call Test_LogOutput(gSL_Select, 1)

    If Not TSMRead(gSL_Select, 0) Then
        Call TSMClose(0)
    Else
        Call GridSet
        DBInput = True
    End If
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
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
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     削除処理を行います      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        .SQL = ""
        .SQL = .SQL & " DELETE FROM 得意先スクラップ単価マスタ"
        .SQL = .SQL & "  WHERE 得意先 = '" & RTrim(iH1_Torcd) & "'"
        If H1Op_Ki(0).Value Then
           .SQL = .SQL & "    AND 適用開始日付 = '" & SY_KamiSMD & "'"
           .SQL = .SQL & "    AND 適用終了日付 = '" & SY_KamiEMD & "'"
        Else
           .SQL = .SQL & "    AND 適用開始日付 = '" & SY_SimoSMD & "'"
           .SQL = .SQL & "    AND 適用終了日付 = '" & SY_SimoEMD & "'"
        End If
        On Error GoTo DBDelete_Err
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        On Error GoTo 0
    End With
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
    Dim AddFlg  As Boolean

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います      "
        .MB_Title = "更新処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
' ( 削除処理 )
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    On Error GoTo DBPut_Err

'------------------------------------'
'   得意先スクラップ単価マスタ出力
'------------------------------------'
    For gInt = 1 To vsGrid1.Rows - 1
        If vsGrid1.TextMatrix(gInt, 0) = "" Then Exit For
        If vsGrid1.TextMatrix(gInt, 3) <> "" Then
            Call ItemsClearTSM
            With TSM
                .得意先 = RTrim(iH1_Torcd)
                .履歴NO = CisFun.Val2(vsGrid1.TextMatrix(gInt, 5)) + 1
                If H1Op_Ki(0).Value Then
                    .適用開始日付 = SY_KamiSMD
                    .適用終了日付 = SY_KamiEMD
                Else
                    .適用開始日付 = SY_SimoSMD
                    .適用終了日付 = SY_SimoEMD
                End If
                .材質 = RTrim(vsGrid1.TextMatrix(gInt, 4))
                .対象年 = RTrim(iH1_Y)
                .対象期 = CisFun.Val2(wKi)
                .単価 = CisFun.Val2(vsGrid1.TextMatrix(gInt, 3))
                .作成者 = gTanto
            End With
            ' ( 追加処理 )
            Call TSMInsert
        End If
    Next gInt
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(VSGRID1)                                 +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設
'定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 20, 1, "材　　質")
        Call .FixedSet(中中, 右中, あり, 12, 2, "前回単価")
        Call .FixedSet(中中, 右中, あり, 12, 3, "今回単価")
        Call .FixedSet(中中, 左中, あり, 0, 4, "材質")
        Call .FixedSet(中中, 右中, あり, 0, 5, "履歴NO")
        
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 6, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
        .InitDisp
    End With
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
    
    If Not TSMRead(gSL_Select) Then Exit Sub
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 > .Rows Then
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not TSM_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid1
            .TextMatrix(ID, GridCol(0)) = ID
            .TextMatrix(ID, GridCol(1)) = RTrim(TSM.材質名)
            .TextMatrix(ID, GridCol(2)) = Format(TSM.単価, "#,##0.00;;#")
            .TextMatrix(ID, GridCol(4)) = RTrim(TSM.材質)
            .TextMatrix(ID, GridCol(5)) = TSM.履歴NO
        End With
        Call TSMReadNext(0)
    Loop

    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call TSMClose(0)
End Sub

'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Area.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = vsGrid1.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid1.Rows - 1 Then
            GoTo Grid_RowChange_End
        Else
            If RTrim(vsGrid1.TextMatrix(gInt + 1, 1)) = "" Then
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
'グリッドからテキストに表示
    Call VsGridToText(vsGrid1.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid1
        '======== 単価 ========
        If CisFun.Val2(iB1_Tan) <> 0 Then
            .TextMatrix(tRow, 3) = Format(CisFun.Val2(iB1_Tan), "#,##0.00")
        Else
            .TextMatrix(tRow, 3) = ""
        End If
                
    End With
    SaveRow = vsGrid1.Row
    vsGrid1.Row = SaveRow
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid1
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, vsGrid1.CellTop, Input_Area.Width, vsGrid1.CellHeight
        '単価
        If CisFun.Val2(.TextMatrix(tRow, 3)) <> 0 Then
            iB1_Tan = Format(CisFun.Val2(.TextMatrix(tRow, 3)), "#,##0.00")
        Else
            iB1_Tan = ""
        End If
        
    End With
End Sub
Private Sub VSGRID1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(vsGrid1.Row)
End Sub
Private Sub VSGRID1_Click()
    Call TextToVsGrid(Val(Input_Area.Tag))
    If vsGrid1.TextMatrix(vsGrid1.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(vsGrid1.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
'*============================*
'*     名称マスタ初期登録     *
'*============================*
Private Sub MeisyoInit()
' 自社情報
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "自社情報"
        .区分桁数 = 1
        .値 = "3"
        .英数字1 = "数字1 = 上期開始月"
        .数字1 = 4
        .英数字2 = "数字2 = 下期開始月"
        .数字2 = 10
        .値名称 = "期"
        .制御区分 = ""
        .作成者 = gTanto
        If Not SYMInitCreate(True) Then
        End If
    End With
' スクラップ単価変動
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "スクラップ単価変動"
        .区分桁数 = 1
        .値 = "1"
        .英数字1 = "数字1 変動率(%)"
        .数字1 = 95
        .値名称 = "率"
        .制御区分 = ""
        .作成者 = gTanto
        If Not SYMInitCreate(True) Then
        End If
    End With
    
    ' 適用月日抽出 & 退避
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & "  WHERE 区分名称 = '自社情報'"
    gSL_Select = gSL_Select & "    AND 値 = '3'"
    gSL_Select = gSL_Select & "    AND 区分タイプ = ''"
    If SYMRead(gSL_Select) Then
       SY_KamiM = Format(SYM.数字1, "00")       ' 上期月
       SY_SimoM = Format(SYM.数字2, "00")       ' 下期月
    End If
End Sub

'+------------------+
'+     日付表示     +
'+------------------+
Private Sub DateDSP(Index As Integer)
    Dim wY As String
    If Len(Trim(iH1_Y)) < 4 And RTrim(iH1_Y) <> "" Then
       If Len(Trim(iH1_Y)) = 1 Then iH1_Y = "200" & iH1_Y     ' 2008.08.29 add
       With CisFun
            .Date_RsFormat = fYMD
            Call .Date_Add(iH1_Y & "0101", D, 1)
            wY = Mid(.Date_Result, 1, 4)
            If Mid(.Date_Result, 1, 1) = "0" Then
                 Mid(wY, 1, 1) = "2"
            End If
            iH1_Y = wY
        End With
    End If
    
    If RTrim(iH1_Y) = "" Then
       H1lb_Kikan = ""
    Else
        If Index = 0 Then
            '***** 上期開始月日抽出 *****
            SY_KamiSMD = iH1_Y & SY_KamiM & "01"
            '***** 上期終了月日抽出 *****
            SY_KamiEMD = iH1_Y & SY_SimoM & "01"
            With CisFun
                .Date_RsFormat = fYMD
                Call .Date_Add(SY_KamiEMD, D, -1)
                SY_KamiEMD = .Date_Result
            End With
            
            If SY_KamiSMD > SY_KamiEMD Then
               SY_KamiEMD = CisFun.Val2(iH1_Y) + 1 & SY_SimoM & "01"
               With CisFun
                   .Date_RsFormat = fYMD
                   Call .Date_Add(SY_KamiEMD, D, -1)
                   SY_KamiEMD = .Date_Result
               End With
            End If
            
            H1lb_Kikan = Mid(SY_KamiSMD, 1, 4) & "/" & _
                         Mid(SY_KamiSMD, 5, 2) & "/" & _
                         Mid(SY_KamiSMD, 7, 2) & "～" & _
                         Mid(SY_KamiEMD, 1, 4) & "/" & _
                         Mid(SY_KamiEMD, 5, 2) & "/" & _
                         Mid(SY_KamiEMD, 7, 2)
        Else
            '***** 下期開始月日抽出 *****
            SY_SimoSMD = iH1_Y & SY_SimoM & "01"
            '***** 下期終了月日抽出 *****
            SY_SimoEMD = iH1_Y & SY_KamiM & "01"
            With CisFun
                .Date_RsFormat = fYMD
                Call .Date_Add(SY_SimoEMD, D, -1)
                SY_SimoEMD = .Date_Result
            End With
            
            If SY_SimoSMD > SY_SimoEMD Then
               SY_SimoEMD = CisFun.Val2(iH1_Y) + 1 & SY_KamiM & "01"
               With CisFun
                   .Date_RsFormat = fYMD
                   Call .Date_Add(SY_SimoEMD, D, -1)
                   SY_SimoEMD = .Date_Result
               End With
            End If
            
            H1lb_Kikan = Mid(SY_SimoSMD, 1, 4) & "/" & _
                         Mid(SY_SimoSMD, 5, 2) & "/" & _
                         Mid(SY_SimoSMD, 7, 2) & "～" & _
                         Mid(SY_SimoEMD, 1, 4) & "/" & _
                         Mid(SY_SimoEMD, 5, 2) & "/" & _
                         Mid(SY_SimoEMD, 7, 2)
        End If
    End If
End Sub

'+------------------+
'+     履歴照会     +
'+------------------+
Private Sub RirekiDisp()
    B2lb_Zai = "材質 ： " & vsGrid1.TextMatrix(vsGrid1.Row, 1)
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT SUBSTRING(適用開始日付, 1, 4) + '/' +"
    gSL_Select = gSL_Select & "        SUBSTRING(適用開始日付, 5, 2) + '/' +"
    gSL_Select = gSL_Select & "        SUBSTRING(適用開始日付, 7, 2) 表示適用開始日付,"
    gSL_Select = gSL_Select & "        SUBSTRING(適用終了日付, 1, 4) + '/' +"
    gSL_Select = gSL_Select & "        SUBSTRING(適用終了日付, 5, 2) + '/' +"
    gSL_Select = gSL_Select & "        SUBSTRING(適用終了日付, 7, 2) 表示適用終了日付,"
    gSL_Select = gSL_Select & " 単価"
    gSL_Select = gSL_Select & "   FROM 得意先スクラップ単価マスタ"
    gSL_Select = gSL_Select & "  WHERE 得意先 = '" & RTrim(iH1_Torcd) & "'"
'    If H1Op_Ki(0).Value Then
'       gSL_Select = gSL_Select & "    AND 適用開始日付 = '" & SY_KamiSMD & "'"
'       gSL_Select = gSL_Select & "    AND 適用終了日付 = '" & SY_KamiEMD & "'"
'    Else
'       gSL_Select = gSL_Select & "    AND 適用開始日付 = '" & SY_SimoSMD & "'"
'       gSL_Select = gSL_Select & "    AND 適用終了日付 = '" & SY_SimoEMD & "'"
'    End If
    gSL_Select = gSL_Select & "    AND 材質 = '" & vsGrid1.TextMatrix(vsGrid1.Row, 4) & "'"
    gSL_Select = gSL_Select & "  ORDER BY 履歴NO DESC"
    
    Adodc1.RecordSource = gSL_Select
    Adodc1.Refresh
    B2lb_Kensu = Adodc1.Recordset.RecordCount & "件"
End Sub
