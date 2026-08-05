VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Begin VB.Form COS0120 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受給照合処理"
   ClientHeight    =   11235
   ClientLeft      =   1050
   ClientTop       =   2955
   ClientWidth     =   15420
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15420
   WindowState     =   2  '最大化
   Begin VB.PictureBox C1_Area1 
      BackColor       =   &H80000007&
      Height          =   3255
      Left            =   11370
      ScaleHeight     =   3195
      ScaleWidth      =   2625
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   2685
      Begin Cis3D_v60.CIS3D UC_3D45 
         Height          =   375
         Index           =   0
         Left            =   30
         Top             =   450
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "照合ＯＫ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Output 
            Height          =   315
            Index           =   1
            Left            =   1710
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   12582912
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D UC_3D49 
         Height          =   375
         Left            =   30
         Top             =   840
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "契約先未検収"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Output 
            Height          =   315
            Index           =   2
            Left            =   1710
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   16711935
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D UC_3D39 
         Height          =   375
         Left            =   30
         Top             =   1620
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "単価差異"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cPositionX      =   30
         cPositionY      =   50
         Begin Cis3D_v60.CIS3D B1lb_Output 
            Height          =   315
            Index           =   4
            Left            =   1710
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   16711935
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D UC_3D42 
         Height          =   375
         Left            =   30
         Top             =   2010
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "数量差異"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Output 
            Height          =   315
            Index           =   5
            Left            =   1710
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   16711935
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   375
         Left            =   30
         Top             =   2400
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "単価数量差異"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Output 
            Height          =   315
            Index           =   6
            Left            =   1710
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   16711935
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   375
         Left            =   30
         Top             =   1230
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "自社未検収"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Output 
            Height          =   315
            Index           =   3
            Left            =   1710
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   16711935
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   405
         Index           =   2
         Left            =   30
         Top             =   30
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   714
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "【 照 合 結 果 】"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   375
         Left            =   30
         Top             =   2790
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "日差異"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Output 
            Height          =   315
            Index           =   7
            Left            =   1710
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   16711935
            Caption         =   "999,999"
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
   Begin VB.PictureBox PR_Area1 
      BackColor       =   &H80000007&
      Height          =   1305
      Left            =   11370
      ScaleHeight     =   1245
      ScaleWidth      =   2625
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3150
      Visible         =   0   'False
      Width           =   2685
      Begin Cis3D_v60.CIS3D UC_3D37 
         Height          =   375
         Index           =   0
         Left            =   30
         Top             =   450
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   661
         BackColor       =   12640511
         ForeColor       =   0
         Caption         =   "契約先売掛金"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Input 
            Height          =   315
            Index           =   1
            Left            =   1650
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   16576
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   375
         Index           =   4
         Left            =   30
         Top             =   840
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   661
         BackColor       =   16777152
         ForeColor       =   0
         Caption         =   "自社受給実績"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Input 
            Height          =   315
            Index           =   2
            Left            =   1650
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   16776960
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   405
         Index           =   1
         Left            =   30
         Top             =   30
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   714
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "【 照合対象件数 】"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   7
      End
   End
   Begin VB.Timer Timer3 
      Left            =   900
      Top             =   1035
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00808080&
      Height          =   4125
      Left            =   4140
      ScaleHeight     =   4065
      ScaleWidth      =   7080
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3150
      Width           =   7140
      Begin VB.PictureBox Picture2 
         Height          =   3915
         Left            =   60
         ScaleHeight     =   3855
         ScaleWidth      =   6900
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   60
         Width           =   6960
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   3720
            Left            =   90
            Top             =   90
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   6562
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin Cis3D_v60.CIS3D UC_3D17 
               Height          =   1575
               Index           =   1
               Left            =   60
               Top             =   45
               Width           =   6645
               _ExtentX        =   11721
               _ExtentY        =   2778
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cPositionX      =   150
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   330
                  Index           =   1
                  Left            =   120
                  Top             =   270
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   582
                  ForeColor       =   16711680
                  Caption         =   "今回照合年月"
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
                  cAlingnment     =   1
                  cBoderStyle     =   2
               End
               Begin CisYM_V60.CisYM iH1_YM 
                  Height          =   420
                  Left            =   1650
                  TabIndex        =   0
                  Top             =   210
                  Width           =   2055
                  _ExtentX        =   2328
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   14.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cSize           =   -1  'True
                  cChkResult      =   0   'False
               End
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   330
                  Index           =   2
                  Left            =   210
                  Top             =   960
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   582
                  ForeColor       =   16711680
                  Caption         =   "契  約  先"
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
                  cAlingnment     =   1
                  cBoderStyle     =   2
               End
               Begin CisText_V60.CisText iH1_Torcd 
                  Height          =   375
                  Left            =   1635
                  TabIndex        =   1
                  Top             =   960
                  Width           =   1020
                  _ExtentX        =   1799
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
               Begin Cis3D_v60.CIS3D H1lb_Torcd 
                  Height          =   360
                  Left            =   3000
                  Top             =   960
                  Width           =   2535
                  _ExtentX        =   4471
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
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
               Begin Cis3D_v60.CIS3D PB_Torcd 
                  Height          =   360
                  Left            =   2655
                  Top             =   960
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   635
                  ForeColor       =   65280
                  Caption         =   "▼"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  cButton         =   -1  'True
                  cPositionX      =   20
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D9 
               Height          =   1995
               Left            =   60
               Top             =   1635
               Width           =   6630
               _ExtentX        =   11695
               _ExtentY        =   3519
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderStyle     =   1
               Begin VB.PictureBox B1_Area 
                  BackColor       =   &H00FFC0FF&
                  Height          =   555
                  Left            =   60
                  ScaleHeight     =   495
                  ScaleWidth      =   6465
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   1380
                  Width           =   6525
                  Begin Cis3D_v60.CIS3D lb_msg 
                     Height          =   435
                     Left            =   30
                     Top             =   30
                     Width           =   5595
                     _ExtentX        =   9869
                     _ExtentY        =   767
                     BackColor       =   16761087
                     Caption         =   "【 処理中 】しばらくお待ち下さい"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   14.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cBoderColor1    =   16761087
                     cBoderColor2    =   16761087
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   7
                     cBoderStyle     =   2
                  End
               End
               Begin VB.PictureBox Picture1 
                  BackColor       =   &H00FFFFC0&
                  Height          =   1305
                  Index           =   1
                  Left            =   480
                  ScaleHeight     =   1245
                  ScaleWidth      =   6045
                  TabIndex        =   6
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   6105
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "( 契約先  )"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   11.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00004080&
                     Height          =   225
                     Index           =   3
                     Left            =   495
                     TabIndex        =   17
                     Top             =   510
                     Width           =   1170
                  End
                  Begin VB.Label H1lb_KTorcd 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXXXX：XXXXXXXXX1XXXXXXXXX2"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   11.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   1830
                     TabIndex        =   16
                     Top             =   510
                     Width           =   4155
                  End
                  Begin VB.Label H1lb_ZYM 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "YYYY 年 MM 月度"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1830
                     TabIndex        =   15
                     Top             =   90
                     Width           =   2025
                  End
                  Begin VB.Label H1lb_ZDate 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "YYYY年 MM月 DD日 HH:NN"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1830
                     TabIndex        =   14
                     Top             =   930
                     Width           =   3075
                  End
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "(対象年月)"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   11.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00004080&
                     Height          =   225
                     Index           =   1
                     Left            =   495
                     TabIndex        =   13
                     Top             =   105
                     Width           =   1140
                  End
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "( 処理日 )"
                     BeginProperty Font 
                        Name            =   "ＭＳ Ｐゴシック"
                        Size            =   11.25
                        Charset         =   128
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00004080&
                     Height          =   225
                     Index           =   0
                     Left            =   525
                     TabIndex        =   12
                     Top             =   945
                     Width           =   1080
                  End
               End
               Begin Cis3D_v60.CIS3D UC_3D10 
                  Height          =   1305
                  Left            =   60
                  Top             =   60
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   2302
                  BackColor       =   0
                  ForeColor       =   65535
                  Caption         =   "前回内容"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   8421504
                  cFont3DColor2   =   -2147483643
                  cFont3D         =   2
                  cAlingnment     =   7
                  cPositionX      =   10
                  cPositionY      =   -290
               End
            End
         End
      End
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   1035
   End
   Begin VB.Timer Timer1 
      Left            =   60
      Top             =   1035
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   75
      TabIndex        =   2
      Text            =   "Dummy"
      Top             =   690
      Width           =   180
   End
   Begin VB.PictureBox Picture7 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   0
   End
   Begin Cis3D_v60.CIS3D PB_ENT 
      Height          =   645
      Left            =   8355
      Top             =   7695
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1138
      ForeColor       =   16576
      Caption         =   "処理開始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
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
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   645
      Left            =   5535
      Top             =   7695
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1138
      ForeColor       =   16711680
      Caption         =   "終  了"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
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
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PB_OK 
      Height          =   645
      Left            =   6945
      Top             =   7695
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1138
      ForeColor       =   65535
      Caption         =   "確  認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
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
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 受 給 照 合 処 理 】"
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
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   11880
         Top             =   30
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
   Begin Cis3D_v60.CIS3D lb_Err 
      Height          =   495
      Left            =   2610
      Top             =   1710
      Visible         =   0   'False
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   873
      BackColor       =   255
      ForeColor       =   16777215
      Caption         =   "Err"
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
      cFont3D         =   2
      cAlingnment     =   7
   End
   Begin VB.Label BG_Logo 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  '透明
      Caption         =   "SIMPR-A"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5535
      TabIndex        =   10
      Top             =   8385
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8340
      TabIndex        =   9
      Top             =   8385
      Width           =   1485
   End
   Begin VB.Label PNL_OK 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6945
      TabIndex        =   8
      Top             =   8385
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "COS0120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   受給照合処理
'**       フォームID    :   COS0120
'**       処理概要      :
'**
'**       作  成  日    :   2004/11/15  By CIS
'**       変  更  日    :   2008/01/24  By CIS - ｷﾉｯｸｽより移行
'**       変  更  日    :   2009/06/10  By CIS - 受給材料対応
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim op_Zairyo               As Byte
    Dim wCnt1                   As Integer
    Dim wCnt2                   As Integer
    
    Dim wRClr                   As Integer
    Dim wGClr                   As Integer
    Dim wBClr                   As Integer
    
    Dim STime                   As Date
    Dim ETime                   As Date
    Dim ProcTime                As Long
    
    Dim mZenYm                  As String
    
    Dim mKTorcd()               As String
    
    Dim mTorcd()                As String
    
    Dim mInCnt(2)               As Long
    Dim mProcCnt(7)             As Long
    
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止  #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # 初 期 内 容 取 得  #
'   #-------------------#
    gConnectCount = 3
    If Not IniGet Then End
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer3.Interval = 300
    Timer3.Enabled = False
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
    
'   < シングルプロセスチェック >
    If Not SingleProcess_Check Then PB_END_Click
    
  ' 名称マスタ確認
    Call MeisyoInit
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   #     ロゴ表示       #
'   #-------------------#
    CisFrm.BGLDisp BG_Logo, 10
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
'   < 前回内容表示 >
    Call ProcRegistry_Check
'   < 対象契約先　取得)
    Call PutKbn
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'   ( グラデーション )
    With CisFrm
        .Gradation_Direction = Portrait
        .Gradation_Shading = Deepen
        .Gradation 100, 100, 150
    End With

End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》               +
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
        Case vbKeyF7:       If Dummy.Tag <> "" Then Call PB_Look_Click      '【検索】
           Case vbKeyF12
                                If PB_END.Visible Then
                                    Call PB_END_Click                           '【終了】
                                Else
                                    If PB_OK.Visible Then Call PB_OK_Click      '【確認】
                                End If
           Case vbKeyReturn
                                If PB_ENT.Visible Then                          '【入力】
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
    
    Call SingleProcess_Release
    Call CisDB.DBDISConnect                         ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'*----------------------------------*
'*      確　認　終　了　キ　ー      *
'*----------------------------------*
Private Sub PB_OK_Click()
    Call PB_END_Click
End Sub
'+----------------------------+
'+      検         索         +
'+----------------------------+
Private Sub PB_Look_Click()
    If Dummy.Tag = "iH1_Torcd" Then Call PB_Torcd_Click: Exit Sub
End Sub
'+---------------------+
'+    契約先検索        +
'+---------------------+
Private Sub PB_Torcd_Click()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "照合対象受給契約先"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
       iH1_Torcd = RV_Code
       H1lb_Torcd = RV_Kubun
    End If

End Sub
'+----------------------------+
'+    iH1_Torcd(GotFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+----------------------------+
'+    iH1_Torcd(LostFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_LostFocus()
    Dummy.Tag = ""
    Call TorGet(iH1_Torcd)
End Sub
Private Function TorGet(Torcd As String) As Boolean
    TorGet = True
    H1lb_Torcd = ""
    If RTrim(Torcd) = "" Then Exit Function
    gSL_Select = "SELECT 値名称,英数字1 "
    gSL_Select = gSL_Select & " FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '照合対象受給契約先'"
    gSL_Select = gSL_Select & "   And 値 = '" & RTrim(Torcd) & "'"
    If SYMRead(gSL_Select, 1) Then
        H1lb_Torcd = SYM.値名称
        H1lb_Torcd.Tag = SYM.英数字1
    Else
        TorGet = False
    End If

End Function
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*       メッセージ表示         *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
End Sub
'*******************************
'*       メッセージ表示         *
'*******************************
Private Sub Timer3_Timer()
'    wCnt1 = wCnt1 + wCnt2
'    If wCnt1 > 7 Then
'        wCnt1 = 6
'        wCnt2 = -1
'    Else
'        If wCnt1 = 0 Then
'            wCnt1 = 2
'            wCnt2 = 1
'        End If
'    End If
'    CIS3D3(0) = String(wCnt1, "▼")
'    CIS3D3(1) = String(8 - wCnt1, "▼")
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "PR"
        Call DispChange(ProcHB)
        DoEvents
        
'       ( Proc Start )
        Call Proc_Start
    
        ProcHB = "C1"
        Call DispChange(ProcHB)
        Timer3.Enabled = True
        
        With CisFrm
            Set .MSGObject = lb_msg
            Set .MSGTimer = Timer2
            .MSGInterval = 100
            .MSGStopCount = 20
            .MSGTYPE = Left_To_Right
            .MSGText = "【 処理終了 】確認して下さい"
            
            .MSGRelate
        End With
        
        GoTo ReturnPress_Ed
    End If
    
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
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
    Dim H1Mode          As Boolean
    Dim H1Color         As Long
    Dim PRMode          As Boolean
    Dim PRColor         As Long
    Dim C1Mode          As Boolean
    Dim C1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    PRMode = False:     PRColor = gPLostSel
    C1Mode = False:     C1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "PR":      PRMode = True:  PRColor = gPGotSel
        Case "C1":      C1Mode = True:  C1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    PR_Area1.Visible = PRMode Or C1Mode
    C1_Area1.Visible = C1Mode
    
    PB_ENT.Visible = Not (PRMode Or C1Mode)
    PNL_ENT.Visible = Not (PRMode Or C1Mode)
    PB_END.Visible = Not (PRMode Or C1Mode)
    PNL_END.Visible = Not (PRMode Or C1Mode)
    
    PB_OK.Visible = C1Mode
    PNL_OK.Visible = C1Mode
    
    Dummy.Enabled = False
End Function
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
'   < 対象年月 >
    If RTrim$(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    今回照合年月を入力して下さい         "
            .MB_Button = Error
            .MBOX
            iH1_YM.SetFocus
            Exit Function
        End With
    End If
    If RTrim$(iH1_YM) < RTrim$(mZenYm) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    前回処理年月より過去の年月は不可！         "
            .MB_Button = Error
            .MBOX
            iH1_YM.SetFocus
            Exit Function
        End With
    End If
    
    If Not TorGet(iH1_Torcd) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    照合対象受給契約先ではありません。       "
            .MB_Button = Error
            .MBOX
            iH1_Torcd.SetFocus
            Exit Function
        End With
    End If
    
    If RTrim(mZenYm) <> "" Then
        gStr = Format(DateAdd("m", 1, CDate(Left(mZenYm, 4) & "/" & Mid(mZenYm, 5, 2) & "/01")), "yyyymm")
        If RTrim$(iH1_YM) > RTrim$(gStr) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "    " & Left(gStr, 4) & " 年 " & Mid(gStr, 5, 2) & " 月度の処理が行われていません！         "
                .MB_MSG(3) = "    　よろしいですか？      "
                .MB_Button = OK_CAN
                If Not .MBOX Then
                    iH1_YM.SetFocus
                    Exit Function
                End If
            End With
        End If
    End If
    
    If Trim(iH1_Torcd) <> "" And Trim(H1lb_Torcd.Tag) = Trim(iH1_YM) Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "    再処理ですが続行しますか？       "
            .MB_MSG(4) = "  OK = 続行、ｷｬﾝｾﾙ = 中止     "
            .MB_Button = OK_CAN
            If Not .MBOX Then
               iH1_YM.SetFocus
               Exit Function
            End If
        End With
    Else
        With CisFun
            .MB_Lines = 4
            gStr = Left(iH1_YM, 4) & " 年 " & Mid(iH1_YM, 5, 2) & "月 "
            .MB_MSG(1) = "   " & gStr & "         "
            .MB_MSG(3) = "    受給照合処理を開始します？         "
            .MB_Button = Yes_No
            If Not .MBOX Then Exit Function
        End With
    End If
    
    ' 処理対象契約先 退避
    ReDim mTorcd(UBound(mKTorcd))
    
    If Trim(iH1_Torcd) <> "" Then
       mTorcd(0) = RTrim(iH1_Torcd)
    Else
       For gInt = 0 To UBound(mKTorcd)
           mTorcd(gInt) = mKTorcd(gInt)
       Next gInt
    End If
    
    For Each gObj In B1lb_Input
        gObj.Caption = ""
    Next gObj
    
    With CisFun
        ' 変換ログ削除日数
        .INI_Section = "受給照合"
        .INI_Default = "120"
        .INI_Key = "LogDelDay"
        .GetIni
        ' 変換ログ出力ファイル設定
        .Log_PutFlg = True
        .Log_FileInitial = "KS"
        .Log_FilePath = gIniExe & "Log\"
        .Log_DataHD = ""
        .Log_DataHDTime = あり
        .Log_DelDay = .Val2(.INI_String)

        .LogPut ""
    
        Call .LogPut("【 受給差異照合 処理開始 】")
    
        Call .LogPut("  （指定内容）　対象年月 = " & iH1_YM)
        Call .LogPut("                契 約 先 = " & iH1_Torcd)
    
    End With
    
    For gCnt1 = 0 To UBound(mTorcd)
        If Trim(mTorcd(gCnt1)) = "" Then Exit For
        If Not SyogoProc(0, mTorcd(gCnt1)) Then Exit Function
    Next gCnt1
    
    ' 契約先買掛金件数
    If CisFun.Val2(B1lb_Input(1)) = 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先売掛金データがありません。       "
            .MB_Button = Error
            .MBOX
            iH1_YM.SetFocus
            Exit Function
        End With
    End If
    ' 自社受給件数
    If CisFun.Val2(B1lb_Input(2)) = 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    自社受給データがありません。       "
            .MB_Button = Error
            .MBOX
            iH1_YM.SetFocus
            Exit Function
        End With
    End If
    
    Head1Chk = True
End Function
'*--------------------------*
'*      処　理　開　始      *
'*--------------------------*
Private Function Proc_Start() As Boolean
    Dim wwShime    As Byte
    
    Proc_Start = False
    
    
    STime = Now()
    
    lb_msg = "【 受給照合処理中 】しばらくお待ち下さい": DoEvents
    
    On Error GoTo Proc_Start_Err
'======================================='
'   【受給照合処理】
'======================================='
    For Each gObj In B1lb_Output
        gObj.Caption = ""
    Next gObj
    For gCnt1 = 0 To UBound(mTorcd)
        If Trim(mTorcd(gCnt1)) = "" Then Exit For
        If Not SyogoProc(1, mTorcd(gCnt1)) Then GoTo Proc_Start_Er
    Next gCnt1


Proc_Start_ED:
    Call CisFun.LogPut("【 受給差異照合 正常終了 】")

'   < 今回内容保存 >
    Call ProcRegistry_Save
    
    On Error GoTo 0
    
    Proc_Start = True
    Exit Function
Proc_Start_Er:

    With CisFrm
        Set .MSGObject = lb_msg
        Set .MSGTimer = Timer2
        .MSGInterval = 100
        .MSGStopCount = 20
        .MSGTYPE = Left_To_Right
        .MSGText = "【 処理異常終了 】"
        
        .MSGRelate
    End With
    Exit Function
Proc_Start_Err:
    CisFun.ErrorBox "受給照合処理"
    Call CisFun.LogPut("【 受給差異照合 異常終了 】")
    End
End Function
'/////////////////////////////////////////////////////
'+---------------------------------------------------+
'+          受給照合処理
'+---------------------------------------------------+
'/////////////////////////////////////////////////////
Private Function SyogoProc(ProcType As Integer, sTorcd As String) As Boolean
    Dim er_No      As Integer
    
    SyogoProc = False
    
    On Error GoTo SyogoProc_ER:
    
    Erase mProcCnt
    Erase mInCnt
    
    Call CisDB.DBTran(TransBegin)
    With CisDB
        .SQL = "受給照合"
        .StoadoCount = 13
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = OutPut用     ' エラー内容
        .ParaNo = 5:    .ParaIO = OutPut用     ' 得意先買掛金入力件数
        .ParaNo = 6:    .ParaIO = OutPut用     ' 自社受給入力件数
        .ParaNo = 7:    .ParaIO = OutPut用     ' 照合ＯＫ件数
        .ParaNo = 8:    .ParaIO = OutPut用     ' 得意先未検収件数
        .ParaNo = 9:    .ParaIO = OutPut用     ' 自社未検収件数
        .ParaNo = 10:    .ParaIO = OutPut用     ' 単価差異件数
        .ParaNo = 11:   .ParaIO = OutPut用     ' 数量差異件数
        .ParaNo = 12:   .ParaIO = OutPut用     ' 単価数量差異件数
        .ParaNo = 13:   .ParaIO = OutPut用     ' 日付差異件数(予備)

        .ParaNo = 1:    .ParaValue = ProcType  ' 処理ﾀｲﾌﾟ
        .ParaNo = 2:    .ParaValue = iH1_YM    ' 対象年月
        .ParaNo = 3:    .ParaValue = sTorcd    ' 対象契約先

        .DBStored

        .ParaNo = 0
        If .ParaValue <> 0 Then
            er_No = .ParaValue
            GoTo SyogoProc_ER
        End If

         .ParaNo = 5: mInCnt(1) = CisFun.Val2(.ParaValue)
         .ParaNo = 6: mInCnt(2) = CisFun.Val2(.ParaValue)
         If ProcType <> 1 Then
            B1lb_Input(1) = Format(CisFun.Val2(B1lb_Input(1)) + CisFun.Val2(mInCnt(1)), "#,###")
            B1lb_Input(2) = Format(CisFun.Val2(B1lb_Input(2)) + CisFun.Val2(mInCnt(2)), "#,###")
         End If
         If ProcType <> 0 Then
            For gInt = 1 To 7
                .ParaNo = gInt + 6
                mProcCnt(gInt) = CisFun.Val2(.ParaValue)
                B1lb_Output(gInt) = Format(CisFun.Val2(B1lb_Output(gInt)) + CisFun.Val2(mProcCnt(gInt)), "#,###")
            Next gInt
         End If
    End With

    With CisFun
        Call .LogPut("受給差異照合 処理結果")
       gStr = "  得意先=" & sTorcd & ",買掛金=" & mInCnt(1) & ",受給=" & mInCnt(2)
       If ProcType <> 0 Then
          gStr = gStr & ",照合OK=" & mProcCnt(1) & ",得意先未検収=" & mProcCnt(2) & ",自社未検収=" & mProcCnt(3)
          gStr = gStr & ",単価差=" & mProcCnt(4) & ",数量差=" & mProcCnt(5) & ",数量･単価差=" & mProcCnt(6)
       End If
       Call .LogPut(gStr)
    End With
    Call CisDB.DBTran(TransCommit)
    
    
    SyogoProc = True
    Exit Function
SyogoProc_ER:
    Select Case er_No
        Case 3: gStr = "受給明細テーブル 再処理用 初期化 エラー"
        Case 4: gStr = "受給明細テーブル 未照合 初期化 エラー"
        Case 5: gStr = "TEXT買掛金テーブル 再処理用 初期化 エラー"
        Case 6: gStr = "TEXT買掛金テーブル 処理対象データ更新 エラー"
        Case 7: gStr = "受給差異明細テーブル削除 エラー"
        Case 8: gStr = "受給明細テーブル照合項目更新 エラー"
        Case 9: gStr = "TEXT買掛金テーブル照合項目更新 エラー"
        Case 10: gStr = "受給明細テーブル照合項目更新 エラー"
        Case 11: gStr = "TEXT買掛金テーブル照合項目更新 エラー"
        Case 20: gStr = "受給照合ワーク作成 エラー"
        Case 21: gStr = "受給照合ワーク金額再計算 エラー"
        Case 31: gStr = " 照合OK 抽出更新 エラー"
        Case 32: gStr = "得意先未検収 抽出更新 エラー"
        Case 33: gStr = "自社未検収 抽出更新 エラー"
        Case 34: gStr = "数量・単価違い 抽出更新 エラー"
        Case 35: gStr = "単価違い 抽出更新 エラー"
        Case 36: gStr = "数量違い 抽出更新 エラー"
        Case 37: gStr = "日付違い 抽出更新 エラー"
        Case 41: gStr = "差異明細テーブル 作成 エラー"
        
        Case 81: gStr = "受給照合ワーク 照合更新 エラー"
        Case 82, 85: gStr = "受給明細テーブル 照合更新 エラー"
        Case 83, 86: gStr = "TEXT買掛金テーブル 照合更新 エラー"
        
        '================================================ 2009/06/10 Start
        Case 13: gStr = "材料購入明細テーブル 再処理用 初期化 エラー"
        Case 14: gStr = "材料購入明細テーブル 未照合 初期化 エラー"
        Case 18: gStr = "材料購入明細テーブル照合項目更新 エラー"
        Case 15: gStr = "材料購入明細テーブル照合項目更新 エラー"
        '================================================ 2009/06/10 End
        
        Case Else
             gStr = "その他ｴﾗｰ:" & Error
    End Select
    
    lb_Err = StrConv(Format(er_No, "00"), vbWide) & ":" & gStr
    
    lb_Err.Visible = True
    
    CisFun.MB_Lines = 5
    CisFun.MB_MSG(2) = "    受給照合処理　異常      "
    CisFun.MB_Button = Error
    CisFun.MBOX
    
    With CisFun
         Call .LogPut(String(10, "-- Err --"))
         Call .LogPut("【 差異照合処理 異常終了 】")
         Call .LogPut(" ")
         Call .LogPut(lb_Err)
         Call .LogPut(" ")
         Call .LogPut(String(10, "-- Err --"))
    End With
    End
    
End Function

'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽
'▽                         名　称　マ　ス　タ　確　認  ( シ ス テ ム 共 通 定 数 等 )
'▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'*---------------------------------------------------------*
'*                  前　回　内　容　獲　得                 *
'*---------------------------------------------------------*
Private Sub ProcRegistry_Check()
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "処理内容"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "受給照合処理"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)

    Call ItemsClearSYM
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '処理内容'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値 = '" & App.EXEName & "'"
    If SYMRead(gSL_Select, 1) Then
        With SYM
            If RTrim$(.英数字1) <> "" Then
                H1lb_ZYM = Format(.英数字1, "@@@@年 @@月")
                mZenYm = RTrim(.英数字1)
            End If
            If RTrim$(.英数字2) <> "" Then
                H1lb_ZDate = .英数字2
            End If
            If RTrim$(.英数字3) <> "" Then
                H1lb_KTorcd = .英数字3
                If TorGet(.英数字3) Then
                    H1lb_KTorcd = H1lb_KTorcd & ":" & H1lb_Torcd
                    H1lb_Torcd = ""
                End If
            End If
        End With
    End If
End Sub
'*---------------------------------------------------------*
'*                  今　回　内　容　保  存                 *
'*---------------------------------------------------------*
Private Sub ProcRegistry_Save()

    Call ItemsClearSYM
    
    With SYM
        .区分名称 = "処理内容"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "受給照合処理"
        .英数字1 = iH1_YM
        .英数字2 = Format(Now(), "YYYY年 MM月 DD日 ") & Format(Now(), "HH:NN")
        .英数字3 = RTrim(iH1_Torcd)
        .作成者 = gTanto
    End With

    Call SYMUpdate

    With CisDB
         .SQL = "Update 名称マスタ set "
         .SQL = .SQL & " 英数字1 = '" & iH1_YM & "'"
         .SQL = .SQL & " Where 区分名称 = '照合対象受給契約先'"
         .SQL = .SQL & "   And 値 <> '' "
         If Trim(iH1_Torcd) <> "" Then
            .SQL = .SQL & " And 値 = '" & iH1_Torcd & "'"
         End If
         .DBExec
    End With

End Sub
'*---------------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス チ ェ ッ ク            *
'*---------------------------------------------------------*
Private Function SingleProcess_Check()
    SingleProcess_Check = False
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "シングルプロセス"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "受給照合処理"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)
    
    Call ItemsClearSYM
    gSL_Select = "Select 英数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称   = 'シングルプロセス'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値        = '" & App.EXEName & "'"
    Call SYMRead(gSL_Select, 1)
    
    With CisFun
        If RTrim$(SYM.英数字1) <> "" And RTrim$(SYM.英数字1) <> .WSID Then
            .MB_Lines = 3
            .MB_MSG(1) = "    他のパソコンで起動中です         "
            .MB_MSG(2) = "        ( " & RTrim$(SYM.英数字1) & " )"
            .MB_Button = Error
            .MBOX
            Exit Function
        End If
    End With

'   < プロセスのロック >
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = SubString( Host_Name(), 1, 20 ),"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

    SingleProcess_Check = True
End Function
'*-------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス 解 除          *
'*-------------------------------------------------*
Private Function SingleProcess_Release()
    SingleProcess_Release = False
    
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = '',"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With
    
    SingleProcess_Release = True
End Function

'+---------------------------------------+
'+   照合対象受給契約先(区分名称ﾏｽﾀ）ﾁｪｯｸ    +
'+---------------------------------------+
Private Sub PutKbn()

    Erase mKTorcd

    gLong = 0
    
    ReDim mKTorcd(1)
    ReDim mTorcd(1)
    
    gSL_Select = "SELECT 値 "
    gSL_Select = gSL_Select & " FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '照合対象受給契約先'"
    gSL_Select = gSL_Select & "   And 値 <> ''"
    If SYMRead(gSL_Select) Then
       ReDim mKTorcd(CisDB.RecordCount)
       ReDim mTorcd(CisDB.RecordCount)
       
       Do Until Not SYM_RDSTS
           
          mKTorcd(gLong) = RTrim(SYM.値)
          gLong = gLong + 1
          
          Call SYMReadNext
       Loop
    End If
    Call SYMClose
End Sub
'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
' 発注区分

    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "照合対象受給契約先"
        .区分桁数 = 7
        .値 = "7364"
        .値名称 = "トヨタ紡織"
        .数字1 = 1
        .作成者 = gTanto
        .検索名 = "照合対象受給契約先 検索"
        .コード名 = "契約先"
        .名称名 = "契約先名"
    End With
    If Not SYMInitCreate(True) Then
    End If
End Sub



