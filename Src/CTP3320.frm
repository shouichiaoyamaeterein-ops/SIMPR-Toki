VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form CTP3320 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "新ｅ－かんばん　前工程情報登録"
   ClientHeight    =   11190
   ClientLeft      =   1845
   ClientTop       =   1005
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
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
      Height          =   1365
      Left            =   1200
      ScaleHeight     =   1305
      ScaleWidth      =   12930
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2175
      Width           =   12990
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   540
         Left            =   6255
         Top             =   720
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   953
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
         Begin CisText_V60.CisText iH1_KeiCD 
            Height          =   360
            Left            =   120
            TabIndex        =   5
            Top             =   105
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
         Begin Cis3D_v60.CIS3D H1lb_KeiCD 
            Height          =   360
            Left            =   1140
            Top             =   105
            Width           =   2505
            _ExtentX        =   4419
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   540
         Left            =   50
         Top             =   725
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   953
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
         Begin Cis3D_v60.CIS3D H1lb_NouJt 
            Height          =   360
            Left            =   3780
            Top             =   105
            Width           =   550
            _ExtentX        =   979
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XX-X"
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
            cAlingnment     =   1
         End
         Begin Cis3D_v60.CIS3D H1lb_NouCD 
            Height          =   360
            Left            =   1170
            Top             =   105
            Width           =   2505
            _ExtentX        =   4419
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
         End
         Begin CisText_V60.CisText iH1_NouCD 
            Height          =   360
            Left            =   120
            TabIndex        =   1
            Top             =   105
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXX"
            MaxLength       =   4
         End
         Begin CisText_V60.CisText iH1_NouKu 
            Height          =   360
            Left            =   765
            TabIndex        =   2
            Top             =   105
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   690
         Index           =   0
         Left            =   50
         Top             =   35
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1217
         ForeColor       =   16711680
         Caption         =   "納  入  先"
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
         cAlingnment     =   7
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   360
            Index           =   14
            Left            =   3645
            Top             =   200
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "事業体"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   2
            cPositionY      =   60
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   690
         Index           =   2
         Left            =   4500
         Top             =   30
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1217
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
         cAlingnment     =   7
         cPositionY      =   60
      End
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   540
         Index           =   3
         Left            =   4500
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   953
         ForeColor       =   16711680
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_NoUke 
            Height          =   375
            Left            =   130
            TabIndex        =   3
            Top             =   105
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   690
         Index           =   4
         Left            =   5120
         Top             =   30
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1217
         ForeColor       =   16711680
         Caption         =   "自工程"
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
         cAlingnment     =   7
         cPositionY      =   60
      End
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   540
         Index           =   5
         Left            =   5120
         Top             =   720
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   953
         ForeColor       =   16711680
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Kote1 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   105
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXX"
            MaxLength       =   6
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   690
         Index           =   7
         Left            =   11775
         Top             =   30
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1217
         ForeColor       =   16711680
         Caption         =   "前工程"
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
         cAlingnment     =   7
         cPositionY      =   60
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   540
         Left            =   10725
         Top             =   725
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   953
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
         Begin Cis3D_v60.CIS3D H1lb_KeiSs 
            Height          =   360
            Left            =   135
            Top             =   105
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXX-X"
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
            cAlingnment     =   1
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   540
         Index           =   9
         Left            =   11775
         Top             =   725
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   953
         ForeColor       =   16711680
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Kote2 
            Height          =   375
            Left            =   105
            TabIndex        =   6
            Top             =   105
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXX"
            MaxLength       =   6
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   540
         Left            =   10035
         Top             =   725
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   953
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
         Begin Cis3D_v60.CIS3D H1lb_KeiJt 
            Height          =   360
            Left            =   180
            Top             =   105
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XX"
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
            cAlingnment     =   1
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   690
         Index           =   10
         Left            =   6255
         Top             =   30
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   1217
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
         cPositionX      =   300
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   360
            Index           =   11
            Left            =   3725
            Top             =   200
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "事業体"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
            cPositionY      =   60
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   360
            Index           =   1
            Left            =   4500
            Top             =   200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "仕入先"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
            cPositionY      =   60
         End
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   165
      Top             =   945
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
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
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   810
      Left            =   8070
      ScaleHeight     =   750
      ScaleWidth      =   6030
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
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
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   -30
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   420
      Width           =   855
   End
   Begin VB.PictureBox XXXXX 
      Height          =   7575
      Left            =   1200
      ScaleHeight     =   7515
      ScaleWidth      =   12930
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2265
      Width           =   12990
      Begin Cis3D_v60.CIS3D B1_Area1 
         Height          =   6000
         Left            =   60
         Top             =   1470
         Width           =   12825
         _ExtentX        =   22622
         _ExtentY        =   10583
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   480
            Left            =   50
            Top             =   60
            Width           =   12735
            _ExtentX        =   22463
            _ExtentY        =   847
            BackColor       =   12648384
            ForeColor       =   16711680
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   20.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionY      =   -120
            Begin Cis3D_v60.CIS3D CIS3D10 
               Height          =   345
               Left            =   11085
               Top             =   45
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   609
               Caption         =   "件 数"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   90
            End
            Begin Cis3D_v60.CIS3D CIS3D8 
               Height          =   330
               Left            =   4290
               Top             =   45
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   582
               BackColor       =   12648384
               ForeColor       =   16711680
               Caption         =   "登 録 内 容"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12648384
               cBoderColor2    =   12648384
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D B1lb_DispSu 
               Height          =   345
               Left            =   11955
               Top             =   45
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   609
               Caption         =   "#,###"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -60
            End
         End
         Begin TrueOleDBGrid60.TDBGrid DBGrid1 
            Bindings        =   "CTP3320.frx":0000
            Height          =   5400
            Left            =   40
            OleObjectBlob   =   "CTP3320.frx":0015
            TabIndex        =   7
            Top             =   520
            Width           =   12720
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   420
   End
   Begin VB.Timer Timer2 
      Left            =   690
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
      Caption         =   "【 新ｅ-かんばん　前工程情報登録  】"
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
         Left            =   10110
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   10
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
               Size            =   11.25
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
                  Size            =   11.25
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
               Size            =   11.25
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
                  Size            =   11.25
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
               Size            =   11.25
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
                  Size            =   11.25
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
         Left            =   7560
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
Attribute VB_Name = "CTP3320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   新ｅ－かんばん　前工程情報登録
'**       フォームID    :   PTP3200
'**       処理概要      :
'**
'**       作  成  日    :   2008/07/15  By CIS SIMPR-A
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim Data_Cnt        As Long         ' 新TOPPS前 工程情報テーブル内のデータ件数
    
    ' 修正用　対象レコードの内容
    Dim Upd_NouSk       As String       ' 納入先
    Dim Upd_NouCD       As String       ' 納入先コード
    Dim Upd_NouKu       As String       ' 納入先工区
    Dim Upd_NouJt       As String       ' 納入先事業体
    Dim Upd_NouJk       As String       ' 納入先事業体工区
    Dim Upd_NoUke       As String       ' 受入
    Dim Upd_Kote1       As String       ' 自工程
    Dim Upd_KeiSk       As String       ' 契約先
    Dim Upd_KeiJt       As String       ' 仕入先事業体
    Dim Upd_KeiSs       As String       ' 仕入先
    Dim Upd_KeiCD       As String       ' 仕入先コード
    Dim Upd_KeiKu       As String       ' 仕入先工区
''  Dim Upd_SyuBa       As String       ' 出荷場
    Dim Upd_Kote2       As String       ' 前工程
    
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()

    '------------------------------ 2005/06/10  Insert
    If Rv_Call3 = "" Then
    '------------------------------ 2005/06/10  Insert End
    '   #------------------#
    '   # 二 重 起 動 防 止 #
    '   #------------------#
        If Not Dupli_Start Then End
    '   #------------------#
    '   # 初 期 内 容 取 得 #
    '   #------------------#
        If Not IniGet Then End
    '------------------------------ 2005/06/10  Insert
    End If
    '------------------------------ 2005/06/10  Insert End
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
    Timer1.Interval = 1000      ' 時刻表示間隔
    Timer1.Enabled = True       ' 時刻表示開始
    Timer2.Interval = 500       ' 処理区分点滅間隔
    Timer2.Enabled = False      ' 処理区分点滅禁止
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
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

    '------------------------------ 2005/06/10  Insert
    If Rv_Call3 = "" Then
    '------------------------------ 2005/06/10  Insert End
       Me.Show
    '------------------------------ 2005/06/10  Insert
    End If
    '------------------------------ 2005/06/10  Insert End
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    FormAct = Not FormAct

' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    Call GridSet                                    ' 2008.07.17 add : 起動時にDB内にあるデータを表示する
    Call DispChange(ProcHB)
    
    '------------------------------ 2005/06/10  Insert
    If Rv_Call3 = "X" Then
'       iH1_Hinbn = RV_Hinbn        ' 2008.07.18 del
'       iH1_Torcd = RV_Torcd        ' 2008.07.18 del
'       Call iH1_Torcd_LostFocus    ' 2008.07.18 del
'       Call_YMD = RV_SYmd          ' 2008.07.18 del
'       CALL_Su = RV_EYmd           ' 2008.07.18 del
'       Back_Call.Visible = True    ' 2008.07.18 del
       
       Call ReturnPress
    End If
    '------------------------------ 2005/06/10  Insert End
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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_LOOK.Visible Then Call PB_Look_Click  '【検索】
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
    '------------------------------ 2005/06/10  Insert
    If Rv_Call3 = "" Then
    '------------------------------ 2005/06/10  Insert End
    '+---------------------+
    '+   ﾃﾞｰﾀﾍﾞｰｽ切断
    '+---------------------+
        Call CisDB.DBDISConnect
    '+---------------------+
    '+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    '+---------------------+
        Call CisFrm.UnLoadDisp
        End
    '------------------------------ 2005/06/10  Insert
    Else
        Me.Hide
    End If
    '------------------------------ 2005/06/10  Insert End
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
      DBGrid1.Row = 0                   ' グリッドのカーソルを最上位に移動
      SyoriKB = "ADD"                   ' B1から取り消したときは新規とする
      Call SyoriName(SyoriKB)
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadClear                    ' H1をクリア
      GoTo PB_CAN_Ed
   End If

   If ProcHB = "H1" Then                ' 修正時、H1から取り消しのときはB1に移動する
      If SyoriKB = "MNT" Then
      ProcHB = "B1"
      Call DispChange(ProcHB)
      Call SyoriName(SyoriKB)
      Call DBGrid1_RowColChange(0, 0)   ' 現在のカーソル位置のデータを再取得する
      End If
   End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    DBGrid1.Row = 0                     ' グリッドのカーソルを最上位にする
    If SyoriKB <> "ADD" Then
       SyoriKB = "ADD"
       Call SyoriName(SyoriKB)
       ProcHB = "H1"
       Call DispChange(ProcHB)
    End If
    Call GridSet
    Call HeadClear                      ' H1をクリア
    iH1_NouCD.SetFocus
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
       SyoriKB = "MNT"
       Call SyoriName(SyoriKB)
       ProcHB = "B1"
       Call DispChange(ProcHB)
       DBGrid1.Row = 0
       Call DBGrid1_RowColChange(0, 0)
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
       SyoriKB = "DEL"
       ProcHB = "B1"
       Call DispChange(ProcHB)
       DBGrid1.Row = 0
       Call DBGrid1_RowColChange(0, 0)
    Else
'        SyoriKB = ""                   ' 2008.07.18 del
    End If
    Call SyoriName(SyoriKB)

End Sub

'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_LOOK.Tag = "iH1_NouCD" Then Call Look_NouCd: Exit Sub
    If PB_LOOK.Tag = "iH1_NouKu" Then Call Look_NouKu: Exit Sub
    If PB_LOOK.Tag = "iH1_NoUke" Then Call Look_NoUke: Exit Sub
    If PB_LOOK.Tag = "iH1_KeiCD" Then Call Look_KeiCd: Exit Sub
End Sub

'+----------------------------+
'+       納入先検索           +
'+----------------------------+
Private Sub Look_NouCd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_Call = ""
   'RV_Kojyo = iH1_NouKu
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_NouCD = Mid(RTrim(RV_Torcd), 1, 4)
        H1lb_NouCD = RV_TorRName
        iH1_NouKu = Mid(RTrim(RV_Kojyo), 1, 1)
    End If
    
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & iH1_NouCD + iH1_NouKu & "'"
    If TRMRead(gSL_Select) Then
        If RTrim(TRM.納入先事業体) <> "" Then
            If RTrim(TRM.納入先事業体工区) <> "" Then
                H1lb_NouJt = TRM.納入先事業体 + "-" + TRM.納入先事業体工区
            Else
                H1lb_NouJt = TRM.納入先事業体
            End If
        End If
    End If
End Sub

'+----------------------------+
'+       納入先工区検索       +
'+----------------------------+
Private Sub Look_NouKu()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_Call = ""
   'RV_Torcd = iH1_NouCD

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_NouCD = Mid(RV_Torcd, 1, 4)
        H1lb_NouCD = RV_TorRName
        iH1_NouKu = Mid(RV_Kojyo, 1, 1)
    End If

    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & iH1_NouCD + iH1_NouKu & "'"
    If TRMRead(gSL_Select) Then
        If RTrim(TRM.納入先事業体) <> "" Then
            If RTrim(TRM.納入先事業体工区) <> "" Then
                H1lb_NouJt = TRM.納入先事業体 + "-" + TRM.納入先事業体工区
            Else
                H1lb_NouJt = TRM.納入先事業体
            End If
        End If
    End If
End Sub

'+----------------------------+
'+       納入先受入検索       +
'+----------------------------+
Private Sub Look_NoUke()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_Call = ""
    RV_Torcd = iH1_NouCD + iH1_NouKu
    RV_Kojyo = iH1_NouKu
'    RV_Call = "C"
'    RV_TorKb = 8
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_NoUke = RV_Ukeir
        iH1_NouCD = RV_TorcdK
        iH1_NouKu = RV_Kojyo
        H1lb_NouCD = RV_TorRName
    End If

    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & iH1_NouCD + iH1_NouKu & "'"
    If TRMRead(gSL_Select) Then
        If RTrim(TRM.納入先事業体) <> "" Then
            If RTrim(TRM.納入先事業体工区) <> "" Then
                H1lb_NouJt = TRM.納入先事業体 + "-" + TRM.納入先事業体工区
            Else
                H1lb_NouJt = TRM.納入先事業体
            End If
        End If
    End If
End Sub

'+----------------------------+
'+       契約先検索           +
'+----------------------------+
Private Sub Look_KeiCd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_KeiCD = RV_TorcdK
        H1lb_KeiCD = RV_TorRName
        
        TRM.取引先CD = RV_TorcdK
        TRM.取引先区分 = 0
   
        gSL_Select = "Select * From 取引先マスタ "
        gSL_Select = gSL_Select & " Where 取引先CD = '" & RTrim(RV_TorcdK) & "'"
        If TRMRead(gSL_Select) Then
            H1lb_KeiJt = TRM.仕入先事業体
            If (RTrim(TRM.被自社)) <> "" Then
                If (RTrim(TRM.被自社工場)) <> "" Then
                    H1lb_KeiSs = Mid(TRM.被自社, 1, 5) + "-" + Mid(TRM.被自社工場, 1, 1)
                Else
                    H1lb_KeiSs = Mid(TRM.被自社, 1, 5)
                End If
            End If
        End If
    End If
End Sub

'+-----------------------------+
'+  iH1_NouCD
'+-----------------------------+
Private Sub iH1_NouCD_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_NouCD_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    H1lb_NouCD = ""
    H1lb_NouJt = ""
    If TorNmGet(iH1_NouCD + iH1_NouKu, 3) Then
'        H1lb_NouCD = iH1_NouCD + "-" + iH1_NouKu + "：" + RTrim(TRM.略称)
'        H1lb_NouJt = RTrim(TRM.納入先事業体 + "-" + TRM.納入先事業体工区)
        H1lb_NouCD = RTrim(TRM.略称)
    End If
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & iH1_NouCD + iH1_NouKu & "'"
    If TRMRead(gSL_Select) Then
        If RTrim(TRM.納入先事業体) <> "" Then
            If RTrim(TRM.納入先事業体工区) <> "" Then
                H1lb_NouJt = TRM.納入先事業体 + "-" + TRM.納入先事業体工区
            Else
                H1lb_NouJt = TRM.納入先事業体
            End If
        End If
    End If
End Sub
'+-----------------------------+
'+  iH1_NouKu
'+-----------------------------+
Private Sub iH1_NouKu_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_NouKu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    H1lb_NouCD = ""
    H1lb_NouJt = ""
    If TorNmGet(iH1_NouCD + iH1_NouKu, 3) Then
'        H1lb_NouCD = iH1_KeiCD + "-" + iH1_NouKu + "：" + RTrim(TRM.略称)
'        H1lb_NouJt = RTrim(TRM.納入先事業体 + "-" + TRM.納入先事業体工区)
        H1lb_NouCD = RTrim(TRM.略称)
    End If
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & iH1_NouCD + iH1_NouKu & "'"
    If TRMRead(gSL_Select) Then
        If RTrim(TRM.納入先事業体) <> "" Then
            If RTrim(TRM.納入先事業体工区) <> "" Then
                H1lb_NouJt = TRM.納入先事業体 + "-" + TRM.納入先事業体工区
            Else
                H1lb_NouJt = TRM.納入先事業体
            End If
        End If
    End If

End Sub
'+-----------------------------+
'+  iH1_NoUke
'+-----------------------------+
Private Sub iH1_NoUke_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_NoUke_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_KeiCD
'+-----------------------------+
Private Sub iH1_KeiCD_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_KeiCD_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    H1lb_KeiCD = ""
    H1lb_KeiJt = ""
    H1lb_KeiSs = ""
    If TorNmGet(iH1_KeiCD, 3) Then
        H1lb_KeiCD = RTrim(TRM.略称)
    End If
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & iH1_KeiCD & "'"
    If TRMRead(gSL_Select) Then
        If RTrim(TRM.仕入先事業体) <> "" Then
            H1lb_KeiJt = TRM.仕入先事業体
        End If
        If (RTrim(TRM.被自社)) <> "" Then
            If (RTrim(TRM.被自社工場)) <> "" Then
                H1lb_KeiSs = Mid(TRM.被自社, 1, 5) + "-" + Mid(TRM.被自社工場, 1, 1)
            Else
                H1lb_KeiSs = Mid(TRM.被自社, 1, 5)
            End If
        End If
    End If
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
     Case "LOK"
          SyoriNM = "照会"
          K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
          PB_ADD.Visible = False
          PB_MNT.Visible = False
          PB_DEL.Visible = False
     Case Else
          SyoriNM = ""
   End Select
   
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
       If SyoriKB <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
          End If
       End If
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
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
    
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
    PB_LOOK.Visible = H1Mode
    PB_ADD.Visible = False: If SyoriKB <> "ADD" Then PB_ADD.Visible = True  ' 2008.07.17 upd : PB_ADD.Visible = H1Mode
    PB_MNT.Visible = False: If SyoriKB <> "MNT" Then PB_MNT.Visible = True  ' 2008.07.17 add : PB_MNT.Visible = H1Mode
    PB_DEL.Visible = False: If SyoriKB <> "DEL" Then PB_DEL.Visible = True  ' 2008.07.17 add : PB_DEL.Visible = H1Mode
    PB_CAN.Visible = False: If SyoriKB <> "ADD" Then PB_CAN.Visible = True
    PB_ENT.Visible = True                                                   ' 2008.07.18 add : PB_ENT.Visible = Not B1Mode
    PB_END.Visible = False: If SyoriKB = "ADD" Then PB_END.Visible = True   ' 2008.07.22 upd : PB_END.Visible = H1Mode
    '------------------------------ 2005/06/10  Insert
    If Data_Cnt = 0 Then
        PB_MNT.Visible = False
        PB_DEL.Visible = False
    End If
    
    If Rv_Call3 <> "" Then
       If Rv_Call3 = "X" Then
          PB_CAN.Visible = False
       End If
       PB_END.BtnCaption = "戻る"
    End If
    '------------------------------ 2005/06/10  Insert End
    
    If Pro = "B1" Then DBGrid1.SetFocus

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
'*---------------------------------- 2008.07.17 del start
'    If HeadBody = "B" Then
'       On Error Resume Next
'       DBGrid1.Close True
'       Adodc1.Recordset.Close
'       Adodc1.RecordSource = ""
'       On Error GoTo 0
'    End If
'*---------------------------------- 2008.07.17 del end
    Call GridSet                    '2008.07.17 add
End Sub
'*---------------------------------- 2008.07.23 add start
'****************************
'*  入力フィールドクリア    *
'****************************
Private Sub HeadClear()
    iH1_NouCD = ""
    iH1_NouKu = ""
    iH1_NoUke = ""
    iH1_Kote1 = ""
    iH1_KeiCD = ""
'   iH1_SyuBa = ""
    iH1_Kote2 = ""

    H1lb_NouCD = ""
    H1lb_NouJt = ""
    H1lb_KeiCD = ""
    H1lb_KeiJt = ""
    H1lb_KeiSs = ""
End Sub
'*---------------------------------- 2008.07.23 add end

'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
    ' << 追加・修正処理 >>
        Call iH1_NouCD_LostFocus        ' 2008.07.24 add : 入力データを再取得する
        Call iH1_NouKu_LostFocus        ' 2008.07.24 add
        Call iH1_NoUke_LostFocus        ' 2008.07.24 add
        Call iH1_KeiCD_LostFocus        ' 2008.07.24 add
        
        If Not Head1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        If SyoriKB = "MNT" Then
            ProcHB = "B1"
            Call DispChange(ProcHB)
        End If
        If SyoriKB = "ADD" Then
            Call HeadClear
            Call DispChange(ProcHB)
        End If
        
        GoTo ReturnPress_Ed
       
        '------------------------------ 2005/06/10  Insert
        If Rv_Call3 = "X" Then
           Rv_Call3Rtn = 1
           Call PB_END_Click
        End If
        '------------------------------ 2005/06/10  Insert End
    End If

    If ProcHB = "B1" Then
        ' << 修正項目選択 >>
        If SyoriKB = "MNT" Then
            Call DBGrid1_DblClick
            GoTo ReturnPress_Ed
        End If
        
        ' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not Head1Chk Then GoTo ReturnPress_Ed
            If Not DBDelete Then GoTo ReturnPress_Ed
            If Data_Cnt = 0 Then
                ProcHB = "H1"                   ' 削除の結果、DB内容が空になったら新規に変更する
                SyoriKB = "ADD"
                Call SyoriName(SyoriKB)
                Call DispChange(ProcHB)
                Call HeadClear
                GoTo ReturnPress_Ed
            End If
        End If
    End If

'    Call HeadBodyClear("B")        ' 2008.07.17 del
'    ProcHB = "H1"                  ' ↓
    Call DispChange(ProcHB)
    Call SyoriName(SyoriKB)

ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub

'****************************
'*      Head1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

' << 項目入力確認 >>
    If Trim(iH1_NouCD) = "" Or Trim(iH1_NouKu) = "" Or Trim(H1lb_NouCD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    納入先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        If Trim(iH1_NouCD) = "" Or Trim(H1lb_NouCD) = "" Then
            iH1_NouCD.SetFocus
        Else
            iH1_NouKu.SetFocus
        End If
        Exit Function
    End If
    If Not TorNmGet(iH1_NouCD + iH1_NouKu, 3) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    納入先未登録            　　　　　"
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_NouCD.SetFocus
        Exit Function
    End If
    If Trim(H1lb_NouJt) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    納入先事業体がありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    If Trim(iH1_NoUke) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    受入を入力して下さい           　 "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_NoUke.SetFocus
        Exit Function
    End If
    
    If Trim(iH1_KeiCD) = "" Or Trim(H1lb_KeiCD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_KeiCD.SetFocus
        Exit Function
    End If
    If Not TorNmGet(iH1_KeiCD, 3) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先未登録           　　　　　 "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_KeiCD.SetFocus
        Exit Function
    End If
    If Trim(H1lb_KeiJt) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先事業体がありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    If Trim(H1lb_KeiSs) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    仕入先がありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
'    If Trim(iH1_Kote1) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    自工程を入力してください      "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'        iH1_Kote1.SetFocus
'        Exit Function
'    End If
'    If Trim(iH1_SyuBa) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    出荷場を入力してください      "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'        iH1_SyuBa.SetFocus
'        Exit Function
'    End If
'    If Trim(iH1_Kote2) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    前工程を入力してください      "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'        iH1_Kote2.SetFocus
'        Exit Function
'    End If

    '< 存在チェック用SQL作成 >
    gSL_Select = "Select * FROM 新TOPPS前工程情報テーブル "
    gSL_Select = gSL_Select & " WHERE 納入先 = '" & RTrim(iH1_NouCD) & RTrim(iH1_NouKu) & "'"
    gSL_Select = gSL_Select & "   And 納入先コード = '" & RTrim(iH1_NouCD) & "'"
    gSL_Select = gSL_Select & "   And 納入先工区   = '" & RTrim(iH1_NouKu) & "'"
    gSL_Select = gSL_Select & "   And 納入先事業体 = '" & Mid(RTrim(H1lb_NouJt), 1, 2) & "'"
    gSL_Select = gSL_Select & "   And 納入先事業体工区 = '" & Mid(RTrim(H1lb_NouJt), 4, 1) & "'"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iH1_NoUke) & "'"
    gSL_Select = gSL_Select & "   And 自工程 = '" & RTrim(iH1_Kote1) & "'"
    gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(iH1_KeiCD) & "'"
    gSL_Select = gSL_Select & "   And 仕入先事業体 = '" & RTrim(H1lb_KeiJt) & "'"
    gSL_Select = gSL_Select & "   And 仕入先コード = '" & Mid(RTrim(H1lb_KeiSs), 1, 4) & "'"
    gSL_Select = gSL_Select & "   And 仕入先工区   = '" & Mid(RTrim(H1lb_KeiSs), 6, 1) & "'"
''  gSL_Select = gSL_Select & "   And 出荷場 = '" & RTrim(iH1_SyuBa) & "'"
    gSL_Select = gSL_Select & "   And 前工程 = '" & RTrim(iH1_Kote2) & "'"

'   << 削除の場合の確認 >>
    If SyoriKB = "DEL" Then
        If Not NTMRead(gSL_Select) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    新TOPPS前工程情報未登録         "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_NouCD.SetFocus
            Exit Function
        End If
    End If

'   << 修正の場合の確認 >>
    If SyoriKB = "MNT" Then
        If (iH1_NouCD <> Upd_NouCD) Or (iH1_NouKu <> Upd_NouKu) Or (iH1_NoUke <> Upd_NoUke) Or _
           (iH1_Kote1 <> Upd_Kote1) Or (iH1_KeiCD <> Upd_KeiSk) Or (iH1_Kote2 <> Upd_Kote2) Then
            If NTMRead(gSL_Select) Then     ' 修正後となるデータが既に存在している場合は修正しない
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    新TOPPS前工程情報登録済         "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_NouCD.SetFocus
                Exit Function
            End If
        End If
        
        If Not UpdDatCheck(0) Then      ' 修正対象データの存在チェックを行なう
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    新TOPPS前工程情報未登録         "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_NouCD.SetFocus
            Exit Function
        End If
    End If
    
'   << 追加の場合の確認 >>
    If SyoriKB = "ADD" Then
        If NTMRead(gSL_Select) Then ' 入力されたデータと同じものが存在するかチェック
        '------------------------------ 2005/06/10  Insert
            If Rv_Call3 = "X" Then
                SyoriKB = "MNT"
                Call SyoriName(SyoriKB)
                Head1Chk = True
                Exit Function
            End If
        '------------------------------ 2005/06/10  Insert End
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    新TOPPS前工程情報登録済         "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_NouCD.SetFocus
            Exit Function
        End If
    End If
    
    Head1Chk = True
End Function

'**************************************
'*                                    *
'*       データベース入力 ＆ 表示     *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    
'   Call HeadBodyClear("B")                 ' 2008.07.17 del
    If Not TRMRead(gSL_Select, 0) Then
        Call TRMClose
    Else
       'Call GridSet
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
End Sub
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
        .MB_MSG(2) = "     削除処理を行います                  "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？                    "
        .MB_Button = OK_CAN
        If Not .MBOX Then GoTo DBDelete_ED 'Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        .SQL = "DELETE FROM 新TOPPS前工程情報テーブル "
        .SQL = .SQL & " WHERE 納入先 = '" & RTrim(iH1_NouCD) & RTrim(iH1_NouKu) & "'"
        .SQL = .SQL & "   And 納入先コード = '" & RTrim(iH1_NouCD) & "'"
        .SQL = .SQL & "   And 納入先工区   = '" & RTrim(iH1_NouKu) & "'"
        .SQL = .SQL & "   And 納入先事業体 = '" & Mid(RTrim(H1lb_NouJt), 1, 2) & "'"
        .SQL = .SQL & "   And 納入先事業体工区 = '" & Mid(RTrim(H1lb_NouJt), 4, 1) & "'"
        .SQL = .SQL & "   And 受入 = '" & RTrim(iH1_NoUke) & "'"
        .SQL = .SQL & "   And 自工程 = '" & RTrim(iH1_Kote1) & "'"
        .SQL = .SQL & "   And 契約先 = '" & RTrim(iH1_KeiCD) & "'"
        .SQL = .SQL & "   And 仕入先事業体 = '" & RTrim(H1lb_KeiJt) & "'"
        .SQL = .SQL & "   And 仕入先コード = '" & Mid(RTrim(H1lb_KeiSs), 1, 4) & "'"
        .SQL = .SQL & "   And 仕入先工区   = '" & Mid(RTrim(H1lb_KeiSs), 6, 1) & "'"
''      .SQL = .SQL & "   And 出荷場 = '" & RTrim(iH1_SyuBa) & "'"
        .SQL = .SQL & "   And 前工程 = '" & RTrim(iH1_Kote2) & "'"
        
        On Error GoTo DBDelete_Err
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
        
        On Error GoTo 0
    End With
    
    Call GridSet            ' DB内容表示を最新の状態に更新する
    Dummy.Enabled = False   ' 2008.07.24 add : 削除後のフォーカスを
    DBGrid1.SetFocus        ' 2008.07.24 add :     DummyからGridにする
    
DBDelete_ED:
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
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "     追加処理を行います                  "
            .MB_Title = "追加処理"
        Else
            .MB_MSG(2) = "     修正処理を行います                  "
            .MB_Title = "修正処理"
        End If
        .MB_MSG(3) = "     よろしいですか？                    "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    If SyoriKB = "MNT" Then         ' 修正の場合は対象データを削除する
        Call UpdDatCheck(1)
    End If
    
    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    On Error GoTo DBPut_Err
    
'   テーブル項目移送
    Call ItemsClearNTM
    With NTM
        .納入先 = RTrim(iH1_NouCD) + RTrim(iH1_NouKu)
        .納入先コード = RTrim(iH1_NouCD)
        .納入先工区 = RTrim(iH1_NouKu)
        .納入先事業体 = Mid(RTrim(H1lb_NouJt), 1, 2)
        .納入先事業体工区 = Mid(RTrim(H1lb_NouJt), 4, 1)
        .受入 = RTrim(iH1_NoUke)
        .自工程 = RTrim(iH1_Kote1)
        .契約先 = RTrim(iH1_KeiCD)
        .仕入先事業体 = RTrim(H1lb_KeiJt)
        .仕入先コード = Mid(RTrim(H1lb_KeiSs), 1, 4)
        .仕入先工区 = Mid(RTrim(H1lb_KeiSs), 6, 1)
        .仕入先 = .仕入先コード + .仕入先工区
''      .出荷場 = RTrim(iH1_SyuBa)
        .前工程 = RTrim(iH1_Kote2)
    End With
    
    ' ( 追加処理 )
    Call NTMInsert              ' DB追加登録
    CisDB.DBTran (TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    If SyoriKB = "ADD" Then
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If
    
    Call GridSet    ' DB内容表示を最新の状態に更新する
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TDBGridPro)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'****************************
'*      内容セット
'****************************
Private Function GridSet() As Boolean
    GridSet = False

    gSL_Select = "Select "
    gSL_Select = gSL_Select & "TB.納入先コード + '-' + TB.納入先工区 + '：' + TR1.略称 編集納入先, "
    gSL_Select = gSL_Select & "TB.納入先事業体 + '-' + TB.納入先事業体工区 編集納入先事業体, "
    gSL_Select = gSL_Select & "TB.受入, TB.自工程, "
    gSL_Select = gSL_Select & "TB.契約先 + '：' + TR2.略称 編集契約先, "
    gSL_Select = gSL_Select & "TB.仕入先事業体, "
    gSL_Select = gSL_Select & "TB.仕入先コード+'-' + TB.仕入先工区 編集仕入先, "
''  gSL_Select = gSL_Select & "TB.出荷場, TB.前工程 "
    gSL_Select = gSL_Select & "TB.前工程 "
    gSL_Select = gSL_Select & "FROM 新TOPPS前工程情報テーブル TB ( NoLock) "
    gSL_Select = gSL_Select & "Left Outer Join 取引先マスタ TR1 ( Nolock ) "
    gSL_Select = gSL_Select & "    ON  TR1.取引先CD   = TB.納入先 "
    gSL_Select = gSL_Select & "    AND TR1.取引先区分 = 0 "
    gSL_Select = gSL_Select & "Left Outer Join 取引先マスタ TR2 ( Nolock ) "
    gSL_Select = gSL_Select & "    ON  TR2.取引先CD   = TB.契約先 "
    gSL_Select = gSL_Select & "    AND TR2.取引先区分 = 0 "
    gSL_Select = gSL_Select & "Order by TB.納入先コード, TB.納入先工区, TB.受入 "
    
    Adodc1.RecordSource = gSL_Select
    Adodc1.Refresh
    B1lb_DispSu = Format(Adodc1.Recordset.RecordCount, "#,##0")
    Data_Cnt = Adodc1.Recordset.RecordCount                         ' Data_Cnt：DB内のデータ件数
    
    If ProcHB = "H1" Then
        iH1_NouCD.SetFocus      ' グリッド表示直後のカーソル位置を指定
        iH1_NouCD_GotFocus      ' グリッド表示直後のカーソル位置を指定
    End If
    
    GridSet = True
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
    If ProcHB = "H1" Then Exit Sub          ' グリッド処理ではないときは処理しない
    If LastRow = "" Then Exit Sub           ' カーソル移動していないときは処理しない
    DoEvents
    Call RtnSet                             ' 移動したカーソル行のデータを取得する
End Sub
'+----------------------------------------------------------+
'+      ADODC1(件数表示)    (不使用)                        +
'+----------------------------------------------------------+
Private Sub TotalDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
End Sub

'+---------------------------------------+
'+  グリッド (Function Key Down)
'+---------------------------------------+
Private Sub DBGrid1_keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then           ' F1キー  (新規:112)
       Call PB_ADD_Click
    End If
    If KeyCode = vbKeyF2 Then           ' F2キー  (修正:113)
       Call PB_MNT_Click
    End If
    If KeyCode = vbKeyF3 Then           ' F3キー  (削除:114)
       Call PB_DEL_Click
    End If
    If KeyCode = vbKeyReturn Then       ' Enterキー   ( 13)
       Call DBGrid1_DblClick
    End If
    If KeyCode = vbKeyEscape Then       ' ESCキー (取消 27)
       Call PB_CAN_Click
    End If
End Sub
''+---------------------------------------+
''+  グリッド (KeyPress)
''+---------------------------------------+
'Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then           ' Enter キー
'       Call DBGrid1_DblClick
'    End If
'End Sub

'+----------------------------------------------------------+
'+      グリッド    ダブルクリック                          +
'+----------------------------------------------------------+
Private Sub DBGrid1_DblClick()
    If ProcHB = "H1" Then GoTo DBGrid1_DblClick_ED
    
    SendKeys "{RIGHT}"
    DoEvents
    Call RtnSet                     ' 選択された行の内容を取得

    If SyoriKB = "MNT" Then
        ProcHB = "H1"               ' データ入力エリアに移動
        Call DispChange(ProcHB)
        GoTo DBGrid1_DblClick_ED
    End If

    If SyoriKB = "DEL" Then
        Call ReturnPress            ' 削除処理を実行する
        GoTo DBGrid1_DblClick_ED
    End If

DBGrid1_DblClick_ED:
End Sub
'+-------------------------------------+
'+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
'+-------------------------------------+
Private Sub RtnSet()
    Dim wk_s As String
    Dim wk_i As Integer

    With Adodc1
'   <修正用　選択データの退避>
        wk_s = RTrim(.Recordset("編集納入先"))
        Upd_NouCD = Mid(wk_s, 1, 4)                             ' 納入先コード
        If Mid(wk_s, 5, 1) = "-" Then                           ' 納入先工区
            Upd_NouKu = Mid(wk_s, 6, 1)
        Else
            Upd_NouKu = ""
        End If
        Upd_NouSk = Upd_NouCD + Upd_NouKu                       ' 納入先
        wk_s = RTrim(.Recordset("編集納入先事業体"))
        Upd_NouJt = Mid(wk_s, 1, 2)                             ' 納入先事業体
        Upd_NouJk = Mid(wk_s, 4, 1)                             ' 納入先事業体工区
        Upd_NoUke = RTrim(.Recordset("受入"))                   ' 受入
        Upd_Kote1 = RTrim(.Recordset("自工程"))                 ' 自工程
        
        wk_s = RTrim(.Recordset("編集契約先"))
        wk_i = InStr(1, wk_s, "：")
        If wk_i <> 0 Then
            Upd_KeiSk = Mid(wk_s, 1, wk_i - 1)                  ' 契約先
        Else
            Upd_KeiSk = ""
        End If
        Upd_KeiJt = RTrim(.Recordset("仕入先事業体"))           ' 仕入先事業体
        wk_s = RTrim(.Recordset("編集仕入先"))
        Upd_KeiCD = Mid(wk_s, 1, 4)                             ' 仕入先コード
        Upd_KeiKu = Mid(wk_s, 6, 1)                             ' 仕入先工区
        Upd_KeiSs = Upd_KeiCD + Upd_KeiKu                       ' 仕入先
        
''      Upd_SyuBa = RTrim(.Recordset("出荷場"))                 ' 出荷場
        Upd_Kote2 = RTrim(.Recordset("前工程"))                 ' 前工程
    
'   <入力＆表示エリアにデータをセット>
        iH1_NouCD = Upd_NouCD                                   ' 納入先CD
        iH1_NouKu = Upd_NouKu                                   ' 納入先工区
        iH1_NoUke = Upd_NoUke                                   ' 受入
        iH1_Kote1 = Upd_Kote1                                   ' 自工程
        iH1_KeiCD = Upd_KeiSk                                   ' 契約先CD
''      iH1_SyuBa = Upd_SyuBa                                   ' 出荷場
        iH1_Kote2 = Upd_Kote2                                   ' 前工程
    
        wk_s = RTrim(.Recordset("編集納入先"))
        H1lb_NouCD = Mid(wk_s, 8, Len(wk_s) - 7)                ' 納品先表示データ
        H1lb_NouJt = RTrim(.Recordset("編集納入先事業体"))      ' 納入先事業体表示データ
        wk_s = RTrim(.Recordset("編集契約先"))
        wk_i = InStr(1, wk_s, "：")
        If wk_i <> 0 Then
            H1lb_KeiCD = Mid(wk_s, wk_i + 1, Len(wk_s) - wk_i)  ' 契約先表示データ
        End If
        H1lb_KeiJt = RTrim(.Recordset("仕入先事業体"))          ' 契約先事業体表示データ
        H1lb_KeiSs = RTrim(.Recordset("編集仕入先"))            ' 仕入先表示データ
    End With
    
    RV_Rtn = True
End Sub

'+--------------------------------------------------------------+
'+  修正データ存在チェック＆削除                                +
'+   (修正データをいったん削除したのち新しいデータを挿入する)   +
'+      flg : 0=チェックのみ／1=チェックと削除                  +
'+--------------------------------------------------------------+
Private Function UpdDatCheck(flg As Integer)
    Dim RunSQL As String
    UpdDatCheck = False
    
    '< 存在チェック用SQL作成 >
    gSL_Select = ""
    gSL_Select = gSL_Select & " WHERE 納入先 = '" & RTrim(Upd_NouCD) & RTrim(Upd_NouKu) & "'"
    gSL_Select = gSL_Select & "   And 納入先コード = '" & RTrim(Upd_NouCD) & "'"
    gSL_Select = gSL_Select & "   And 納入先工区   = '" & RTrim(Upd_NouKu) & "'"
    gSL_Select = gSL_Select & "   And 納入先事業体 = '" & RTrim(Upd_NouJt) & "'"
    gSL_Select = gSL_Select & "   And 納入先事業体工区 = '" & RTrim(Upd_NouJk) & "'"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(Upd_NoUke) & "'"
    gSL_Select = gSL_Select & "   And 自工程 = '" & RTrim(Upd_Kote1) & "'"
    gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(Upd_KeiSk) & "'"
    gSL_Select = gSL_Select & "   And 仕入先事業体 = '" & RTrim(Upd_KeiJt) & "'"
    gSL_Select = gSL_Select & "   And 仕入先 = '" & RTrim(Upd_KeiSs) & "'"
    gSL_Select = gSL_Select & "   And 仕入先コード = '" & RTrim(Upd_KeiCD) & "'"
    gSL_Select = gSL_Select & "   And 仕入先工区   = '" & RTrim(Upd_KeiKu) & "'"
''  gSL_Select = gSL_Select & "   And 出荷場 = '" & RTrim(Upd_SyuBa) & "'"
    gSL_Select = gSL_Select & "   And 前工程 = '" & RTrim(Upd_Kote2) & "'"
    
    RunSQL = "Select * FROM 新TOPPS前工程情報テーブル " & gSL_Select
    
    If Not NTMRead(RunSQL) Then Exit Function           ' 存在しないデータだったら中断
    
    If flg <> 0 Then    ' チェックと削除の指示ならば修正対象データを削除する
        RunSQL = "Delete FROM 新TOPPS前工程情報テーブル " + gSL_Select
        ' 既存のデータを削除する
        With CisDB
            On Error GoTo UpdDatCheck_Err
            .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
            .SQL = RunSQL
            .DBExec
            .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        End With
    End If
    
    UpdDatCheck = True
    Exit Function
UpdDatCheck_Err:
    CisFun.ErrorBox
    End
End Function
