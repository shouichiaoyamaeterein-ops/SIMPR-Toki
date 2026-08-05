VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Begin VB.Form COK0110 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "買掛実績作成処理"
   ClientHeight    =   11235
   ClientLeft      =   -2760
   ClientTop       =   1095
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
   Begin VB.Timer Timer3 
      Left            =   900
      Top             =   1035
   End
   Begin VB.PictureBox C1_Area1 
      BackColor       =   &H80000007&
      Height          =   10470
      Left            =   10575
      ScaleHeight     =   10410
      ScaleWidth      =   4035
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   4095
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   0
         Left            =   45
         Top             =   2685
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "OUT:購入実績件数"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   1
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   405
         Index           =   0
         Left            =   45
         Top             =   30
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   714
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "<<<　処 理 対 象 件 数　>>>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   1
         Left            =   45
         Top             =   570
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "IN:購入データ"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   1
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   3
         Left            =   45
         Top             =   3210
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "OUT:単価未決件数"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   2
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   2
         Left            =   45
         Top             =   1095
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "IN:型費データ"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   2
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   4
         Left            =   45
         Top             =   4890
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "OUT:支給実績件数"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   0
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   5
         Left            =   45
         Top             =   3825
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "IN:支給データ"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   0
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   6
         Left            =   45
         Top             =   5415
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "OUT:単価未決件数"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   4
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   7
         Left            =   45
         Top             =   9840
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "OUT:請求書件数"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   3
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   8
         Left            =   45
         Top             =   1620
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "IN:購入訂正"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   3
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   9
         Left            =   45
         Top             =   4350
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "IN:支給訂正"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   4
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   10
         Left            =   45
         Top             =   2145
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "IN:受給データ"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   5
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   405
         Index           =   1
         Left            =   45
         Top             =   6045
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   714
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "<<<　材料処理対象件数　>>>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   11
         Left            =   45
         Top             =   6465
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "IN:購入データ"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   6
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   12
         Left            =   45
         Top             =   8085
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "IN:支給データ"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   360
            Index           =   7
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   33023
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   13
         Left            =   45
         Top             =   6990
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "OUT:購入実績件数"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   5
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   14
         Left            =   45
         Top             =   7515
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "OUT:単価未決件数"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   6
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   15
         Left            =   45
         Top             =   8610
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "OUT:支給実績件数"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   7
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   525
         Index           =   16
         Left            =   45
         Top             =   9135
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   926
         BackColor       =   12648447
         ForeColor       =   0
         Caption         =   "OUT:単価未決件数"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D C1lb_Output 
            Height          =   360
            Index           =   8
            Left            =   1995
            Top             =   90
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   0
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
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00808080&
      Height          =   5235
      Left            =   4410
      ScaleHeight     =   5175
      ScaleWidth      =   7080
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2880
      Width           =   7140
      Begin VB.PictureBox Picture2 
         Height          =   5025
         Left            =   60
         ScaleHeight     =   4965
         ScaleWidth      =   6900
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   60
         Width           =   6960
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   4830
            Left            =   90
            Top             =   90
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   8520
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
               Height          =   2415
               Index           =   1
               Left            =   60
               Top             =   45
               Width           =   6645
               _ExtentX        =   11721
               _ExtentY        =   4260
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
               Begin Cis3D_v60.CIS3D CIS3D2 
                  Height          =   555
                  Left            =   1650
                  Top             =   750
                  Width           =   4875
                  _ExtentX        =   8599
                  _ExtentY        =   979
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderStyle     =   1
                  Begin CisText_V60.CisText iH1_Shime 
                     Height          =   360
                     Left            =   1860
                     TabIndex        =   4
                     Top             =   120
                     Width           =   420
                     _ExtentX        =   741
                     _ExtentY        =   635
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
                     cFdAutoFormat   =   1
                     cGFormat        =   "##"
                     cILength        =   2
                  End
                  Begin VB.OptionButton H1Op_Shime 
                     Caption         =   "個別(実績更新済み)を除く全て"
                     Height          =   300
                     Index           =   2
                     Left            =   1035
                     TabIndex        =   3
                     Top             =   660
                     Visible         =   0   'False
                     Width           =   3810
                  End
                  Begin VB.OptionButton H1Op_Shime 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "個別"
                     Height          =   300
                     Index           =   1
                     Left            =   1035
                     TabIndex        =   2
                     Top             =   150
                     Width           =   840
                  End
                  Begin VB.OptionButton H1Op_Shime 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "全て"
                     Height          =   300
                     Index           =   0
                     Left            =   120
                     TabIndex        =   1
                     Top             =   150
                     Width           =   840
                  End
                  Begin VB.Label Label1 
                     Caption         =   "末日は'99'にて指定"
                     Height          =   255
                     Left            =   2325
                     TabIndex        =   27
                     Top             =   180
                     Width           =   2325
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   330
                  Index           =   1
                  Left            =   120
                  Top             =   270
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   582
                  ForeColor       =   16711680
                  Caption         =   "今回対象年月"
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
                  Height          =   435
                  Left            =   1650
                  TabIndex        =   0
                  Top             =   210
                  Width           =   2025
                  _ExtentX        =   2646
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
                  cContent        =   1
               End
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   330
                  Index           =   0
                  Left            =   210
                  Top             =   885
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   582
                  ForeColor       =   16711680
                  Caption         =   "締      日"
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
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   330
                  Index           =   2
                  Left            =   210
                  Top             =   1440
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
                  TabIndex        =   5
                  Top             =   1440
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
                  Top             =   1440
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
                  Top             =   1440
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
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   330
                  Index           =   3
                  Left            =   165
                  Top             =   1950
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   582
                  ForeColor       =   16711680
                  Caption         =   "処 理 区 分"
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
               Begin Cis3D_v60.CIS3D CIS3D4 
                  Height          =   405
                  Left            =   1635
                  Top             =   1905
                  Width           =   4875
                  _ExtentX        =   8599
                  _ExtentY        =   714
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderStyle     =   1
                  Begin VB.OptionButton H1Op_Proc 
                     Caption         =   "新規"
                     Height          =   300
                     Index           =   0
                     Left            =   120
                     TabIndex        =   6
                     Top             =   60
                     Width           =   840
                  End
                  Begin VB.OptionButton H1Op_Proc 
                     Caption         =   "再処理"
                     Height          =   300
                     Index           =   1
                     Left            =   1470
                     TabIndex        =   7
                     Top             =   60
                     Width           =   1170
                  End
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D9 
               Height          =   2265
               Left            =   60
               Top             =   2505
               Width           =   6630
               _ExtentX        =   11695
               _ExtentY        =   3995
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
                  TabIndex        =   14
                  TabStop         =   0   'False
                  Top             =   1650
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
                  Height          =   1545
                  Index           =   1
                  Left            =   480
                  ScaleHeight     =   1485
                  ScaleWidth      =   6045
                  TabIndex        =   13
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   6105
                  Begin VB.Label H1lb_Proc 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "新規"
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
                     Left            =   1845
                     TabIndex        =   29
                     Top             =   915
                     Width           =   480
                  End
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "(処理区分)"
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
                     Index           =   4
                     Left            =   525
                     TabIndex        =   28
                     Top             =   915
                     Width           =   1140
                  End
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
                     TabIndex        =   26
                     Top             =   630
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
                     TabIndex        =   25
                     Top             =   630
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
                     TabIndex        =   24
                     Top             =   90
                     Width           =   2025
                  End
                  Begin VB.Label H1lb_Shime 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXXXXXXXXXXXXXXX"
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
                     TabIndex        =   23
                     Top             =   345
                     Width           =   2970
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
                     TabIndex        =   22
                     Top             =   1200
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
                     TabIndex        =   21
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
                     TabIndex        =   20
                     Top             =   1215
                     Width           =   1080
                  End
                  Begin VB.Label Label2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "(　締　日　)"
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
                     Index           =   2
                     Left            =   495
                     TabIndex        =   19
                     Top             =   360
                     Width           =   1155
                  End
               End
               Begin Cis3D_v60.CIS3D UC_3D10 
                  Height          =   1545
                  Left            =   60
                  Top             =   60
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   2725
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
      TabIndex        =   8
      Text            =   "Dummy"
      Top             =   690
      Width           =   180
   End
   Begin VB.PictureBox Picture7 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   0
   End
   Begin Cis3D_v60.CIS3D PB_ENT 
      Height          =   645
      Left            =   8355
      Top             =   8205
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
      Top             =   8205
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
      Top             =   8205
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
      Caption         =   "【 買 掛 実 績 作 成 処 理 】"
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   8895
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8340
      TabIndex        =   16
      Top             =   8895
      Width           =   1485
   End
   Begin VB.Label PNL_OK 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6945
      TabIndex        =   15
      Top             =   8895
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "COK0110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   買掛実績作成処理
'**       フォームID    :   COK0110
'**       処理概要      :
'**
'**       作  成  日    :   2004/11/25  By CIS
'**       変  更  日    :   2008/01/22  受給実績を追加
'**       変  更  日    :   2008/06/13  材料実績追加
'**
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
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long
    
    Dim mShime                  As Byte
    Dim mTorcd                  As String
    Dim mSyori                  As String
    Dim mZenYm                  As String
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
    
    mGotColor = H1Op_Shime(0).BackColor
    mLostColor = H1Op_Shime(1).BackColor
'   < シングルプロセスチェック >
    If Not SingleProcess_Check Then PB_END_Click
    
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
    
    H1Op_Shime(0).Value = True
    H1Op_Proc(0).Value = True
    
'   < 前回内容表示 >
    Call ProcRegistry_Check
'   < 名称マスタ初期設定 >
    Call MeisyoInit
    
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
        .Gradation 120, 200, 100
    End With

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
    
    Call KTanErr_List
    Call STanErr_List

'====================================== *** 2008/06/13 Start
    Call ZKTanErr_List
    Call ZSTanErr_List
'====================================== *** 2008/06/13 End

    Call SingleProcess_Release
    Call CisDB.DBDISConnect                         ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    End
End Sub
'*-------------------------------------------------*
'*      購入単価未決リスト
'*-------------------------------------------------*
Public Sub KTanErr_List()
    
'   < 印刷データ読み込み >
    gSL_Select = "SELECT HN.表示品番,TW.契約先," & vbCr
    gSL_Select = gSL_Select & " TW.品番,TW.仕入先," & vbCr
    gSL_Select = gSL_Select & " TW.検収日,SUM(TW.数量) 数量," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR1.略称,'') 仕入先名," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR2.略称,'') 契約先名" & vbCr
    gSL_Select = gSL_Select & " From 購入単価未決ワーク TW" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN" & vbCr
    gSL_Select = gSL_Select & "    ON HN.品番 = TW.品番" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1" & vbCr
    gSL_Select = gSL_Select & "    ON TR1.取引先CD = TW.仕入先" & vbCr
    gSL_Select = gSL_Select & "    And TR1.取引先区分 = 1" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2" & vbCr
    gSL_Select = gSL_Select & "    ON TR2.取引先CD = TW.契約先" & vbCr
    gSL_Select = gSL_Select & "    And TR2.取引先区分 = 1" & vbCr

    gSL_Select = gSL_Select & " GROUP BY HN.表示品番,TW.契約先,TW.品番,TW.仕入先,TW.検収日,TR1.略称,TR2.略称" & vbCr
    gSL_Select = gSL_Select & " Order By TW.仕入先,HN.表示品番,TW.契約先,TW.検収日" & vbCr
    If Not KJSRead(gSL_Select) Then
       Call KJSClose
       Exit Sub
    End If
    
    Dim CisSimplePrint As New CisSimplePrint
    
    With CisSimplePrint
        .PrintTitle = "購入単価未決リスト"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Portrait
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        
        .pHeadItemSet 0, 中, "仕  入  先"
        .pHeadItemSet 1, 中, "品  番"
        .pHeadItemSet 2, 中, "検 収 日"
        .pHeadItemSet 3, 右, "検収数"
        .pHeadItemSet 4, 中, "契　約　先"
    
        .pBodyItemSet 0, 20
        .pBodyItemSet 1, 25
        .pBodyItemSet 2, 10
        .pBodyItemSet 3, 8, 右
        .pBodyItemSet 4, 25
        
        .BodyInit
        .MaxCount = CisDB.RecordCount
        
        If Not .CheckPrtCsv Then Exit Sub
        
        Do Until (Not KJS_RDSTS)
            
            gStr = KJS.仕入先
            If Trim(KJS.仕入先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(KJS.仕入先名, 1, 20)
            .pBodyValueSet 0, gStr
            If Trim(KJS.表示品番) = "" Then
               .pBodyValueSet 1, RTrim$(KJS.品番)
            Else
               .pBodyValueSet 1, RTrim$(KJS.表示品番)
            End If
            With CisFun
                .DateE_BefVal = KJS.検収日
                .DateEdit
            End With
            .pBodyValueSet 2, CisFun.DateE_AfVal
            .pBodyValueSet 3, Format(KJS.数量, "#,###")
            gStr = KJS.契約先
            If Trim(KJS.契約先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(KJS.契約先名, 1, 20)
            .pBodyValueSet 4, gStr
            
            .BodySet
            
            Call KJSReadNext
        Loop
        
        .PrintStart
    
    End With
    
End Sub
'*-------------------------------------------------*
'*      支給単価未決リスト
'*-------------------------------------------------*
Public Sub STanErr_List()
    
'   < 印刷データ読み込み >
    gSL_Select = "SELECT HN.表示品番,TW.契約先," & vbCr
    gSL_Select = gSL_Select & " TW.品番,TW.支給先," & vbCr
    gSL_Select = gSL_Select & " TW.検収日,SUM(TW.数量) 数量," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR1.略称,'') 支給先名," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR2.略称,'') 契約先名" & vbCr
    gSL_Select = gSL_Select & " From 支給単価未決ワーク TW" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN" & vbCr
    gSL_Select = gSL_Select & "    ON HN.品番 = TW.品番" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1" & vbCr
    gSL_Select = gSL_Select & "    ON TR1.取引先CD = TW.支給先" & vbCr
    gSL_Select = gSL_Select & "    And TR1.取引先区分 = 1" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2" & vbCr
    gSL_Select = gSL_Select & "    ON TR2.取引先CD = TW.契約先" & vbCr
    gSL_Select = gSL_Select & "    And TR2.取引先区分 = 1" & vbCr

    gSL_Select = gSL_Select & " GROUP BY HN.表示品番,TW.契約先,TW.品番,TW.支給先,TW.検収日,TR1.略称,TR2.略称" & vbCr
    gSL_Select = gSL_Select & " Order By TW.支給先,HN.表示品番,TW.契約先,TW.検収日" & vbCr
    If Not SJSRead(gSL_Select) Then
       Call SJSClose
       Exit Sub
    End If
    
    Dim CisSimplePrint As New CisSimplePrint
    
    With CisSimplePrint
        .PrintTitle = "支給単価未決リスト"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Portrait
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        
        .pHeadItemSet 0, 中, "支  給  先"
        .pHeadItemSet 1, 中, "品  番"
        .pHeadItemSet 2, 中, "検 収 日"
        .pHeadItemSet 3, 右, "検収数"
        .pHeadItemSet 4, 中, "契　約　先"
    
        .pBodyItemSet 0, 20
        .pBodyItemSet 1, 25
        .pBodyItemSet 2, 10
        .pBodyItemSet 3, 8, 右
        .pBodyItemSet 4, 25
        
        .BodyInit
        .MaxCount = CisDB.RecordCount
        
        If Not .CheckPrtCsv Then Exit Sub
        
        Do Until (Not SJS_RDSTS)
            
            gStr = SJS.支給先
            If Trim(SJS.支給先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(SJS.支給先名, 1, 20)
            .pBodyValueSet 0, gStr
            If Trim(SJS.表示品番) = "" Then
               .pBodyValueSet 1, RTrim$(SJS.品番)
            Else
               .pBodyValueSet 1, RTrim$(SJS.表示品番)
            End If
            With CisFun
                .DateE_BefVal = SJS.検収日
                .DateEdit
            End With
            .pBodyValueSet 2, CisFun.DateE_AfVal
            .pBodyValueSet 3, Format(SJS.数量, "#,###")
            gStr = SJS.契約先
            If Trim(SJS.契約先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(SJS.契約先名, 1, 20)
            .pBodyValueSet 4, gStr
            
            .BodySet
            
            Call SJSReadNext
        Loop
        
        .PrintStart
    
    End With
    
End Sub

'==================================================================== *** 2008/06/13 Start
'*-------------------------------------------------*
'*      材料購入単価未決リスト
'*-------------------------------------------------*
Public Sub ZKTanErr_List()
    
'   < 印刷データ読み込み >
    gSL_Select = "SELECT ZN.材質,ZN.表示寸法,TW.契約先," & vbCr
    gSL_Select = gSL_Select & " TW.材料管理番号,TW.仕入先," & vbCr
    gSL_Select = gSL_Select & " TW.検収日,SUM(TW.数量) 数量," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR1.略称,'') 仕入先名," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR2.略称,'') 契約先名" & vbCr
    gSL_Select = gSL_Select & " From 材料購入単価未決ワーク TW" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ ZN" & vbCr
    gSL_Select = gSL_Select & "    ON ZN.材料管理番号 = TW.材料管理番号" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1" & vbCr
    gSL_Select = gSL_Select & "    ON TR1.取引先CD = TW.仕入先" & vbCr
    gSL_Select = gSL_Select & "    And TR1.取引先区分 = 1" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2" & vbCr
    gSL_Select = gSL_Select & "    ON TR2.取引先CD = TW.契約先" & vbCr
    gSL_Select = gSL_Select & "    And TR2.取引先区分 = 1" & vbCr

    gSL_Select = gSL_Select & " GROUP BY ZN.材質,ZN.表示寸法,TW.契約先,TW.材料管理番号,TW.仕入先,TW.検収日,TR1.略称,TR2.略称" & vbCr
    gSL_Select = gSL_Select & " Order By TW.仕入先,ZN.材質,ZN.表示寸法,TW.契約先,TW.検収日" & vbCr
    If Not ZKJRead(gSL_Select) Then
       Call ZKJClose
       Exit Sub
    End If
    
    Dim CisSimplePrint As New CisSimplePrint
    
    With CisSimplePrint
        .PrintTitle = "材料購入単価未決リスト"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Portrait
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        
        .pHeadItemSet 0, 中, "仕  入  先"
        .pHeadItemSet 1, 中, "材  質"
        .pHeadItemSet 2, 中, "寸      法"
        .pHeadItemSet 3, 中, "検 収 日"
        .pHeadItemSet 4, 右, "検収数"
        .pHeadItemSet 5, 中, "契約先"
    
        .pBodyItemSet 0, 15
        .pBodyItemSet 1, 23
        .pBodyItemSet 2, 25
        .pBodyItemSet 3, 10
        .pBodyItemSet 4, 8, 右
        .pBodyItemSet 5, 8
        
        .BodyInit
        .MaxCount = CisDB.RecordCount
        
        If Not .CheckPrtCsv Then Exit Sub
        
        Do Until (Not ZKJ_RDSTS)
            
            gStr = ZKJ.仕入先
            If Trim(ZKJ.仕入先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(ZKJ.仕入先名, 1, 20)
            .pBodyValueSet 0, gStr
            .pBodyValueSet 1, RTrim$(ZKJ.材質)
            .pBodyValueSet 2, RTrim$(ZKJ.表示寸法)
            
            With CisFun
                .DateE_BefVal = ZKJ.検収日
                .DateEdit
            End With
            .pBodyValueSet 3, CisFun.DateE_AfVal
            .pBodyValueSet 4, Format(ZKJ.数量, "#,###")
            gStr = ZKJ.契約先
            'If Trim(ZKJ.契約先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(ZKJ.契約先名, 1, 20)
            .pBodyValueSet 5, gStr
            
            .BodySet
            
            Call ZKJReadNext
        Loop
        
        .PrintStart
    
    End With
    
End Sub
'*-------------------------------------------------*
'*      材料支給単価未決リスト
'*-------------------------------------------------*
Public Sub ZSTanErr_List()
    
'   < 印刷データ読み込み >
    gSL_Select = "SELECT ZN.材質,ZN.表示寸法,TW.契約先," & vbCr
    gSL_Select = gSL_Select & " TW.材料管理番号,TW.支給先," & vbCr
    gSL_Select = gSL_Select & " TW.検収日,SUM(TW.数量) 数量," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR1.略称,'') 支給先名," & vbCr
    gSL_Select = gSL_Select & " ISNULL(TR2.略称,'') 契約先名" & vbCr
    gSL_Select = gSL_Select & " From 材料支給単価未決ワーク TW" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ ZN" & vbCr
    gSL_Select = gSL_Select & "    ON ZN.材料管理番号 = TW.材料管理番号" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1" & vbCr
    gSL_Select = gSL_Select & "    ON TR1.取引先CD = TW.支給先" & vbCr
    gSL_Select = gSL_Select & "    And TR1.取引先区分 = 1" & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2" & vbCr
    gSL_Select = gSL_Select & "    ON TR2.取引先CD = TW.契約先" & vbCr
    gSL_Select = gSL_Select & "    And TR2.取引先区分 = 1" & vbCr

    gSL_Select = gSL_Select & " GROUP BY ZN.材質,ZN.表示寸法,TW.契約先,TW.材料管理番号,TW.支給先,TW.検収日,TR1.略称,TR2.略称" & vbCr
    gSL_Select = gSL_Select & " Order By TW.支給先,ZN.材質,ZN.表示寸法,TW.契約先,TW.検収日" & vbCr
    If Not ZSJRead(gSL_Select) Then
       Call ZSJClose
       Exit Sub
    End If
    
    Dim CisSimplePrint As New CisSimplePrint
    
    With CisSimplePrint
        .PrintTitle = "材料支給単価未決リスト"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Portrait
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        
        .pHeadItemSet 0, 中, "支  給  先"
        .pHeadItemSet 1, 中, "材  質"
        .pHeadItemSet 2, 中, "寸      法"
        .pHeadItemSet 3, 中, "検 収 日"
        .pHeadItemSet 4, 右, "検収数"
        .pHeadItemSet 5, 中, "契約先"
    
        .pBodyItemSet 0, 15
        .pBodyItemSet 1, 23
        .pBodyItemSet 2, 25
        .pBodyItemSet 3, 10
        .pBodyItemSet 4, 8, 右
        .pBodyItemSet 5, 8
        
        .BodyInit
        .MaxCount = CisDB.RecordCount
        
        If Not .CheckPrtCsv Then Exit Sub
        
        Do Until (Not ZSJ_RDSTS)
            
            gStr = ZSJ.支給先
            If Trim(ZSJ.支給先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(ZSJ.支給先名, 1, 20)
            .pBodyValueSet 0, gStr
            .pBodyValueSet 1, RTrim$(ZSJ.材質)
            .pBodyValueSet 2, RTrim$(ZSJ.表示寸法)
            With CisFun
                .DateE_BefVal = ZSJ.検収日
                .DateEdit
            End With
            .pBodyValueSet 3, CisFun.DateE_AfVal
            .pBodyValueSet 4, Format(ZSJ.数量, "#,###")
            gStr = ZSJ.契約先
            'If Trim(ZSJ.契約先名) <> "" Then gStr = gStr & "：" & CisFun.Mid2(ZSJ.契約先名, 1, 20)
            .pBodyValueSet 5, gStr
            
            .BodySet
            
            Call ZSJReadNext
        Loop
        
        .PrintStart
    
    End With
    
End Sub
'==================================================================== *** 2008/06/13 End

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
'+----------------------------------------+
'+  H1Op_Shime   締日
'+----------------------------------------+
Private Sub H1Op_Shime_Click(Index As Integer)
    
    H1Op_Shime(0).Tag = Index
    For Each gObj In H1Op_Shime
        If gObj.Index = Index Then
            gObj.BackColor = mGotColor
        Else
            gObj.BackColor = mLostColor
        End If
    Next gObj
    If Index = 0 Then
       iH1_Shime.Enabled = False
       iH1_Shime = ""
    Else
       iH1_Shime.Enabled = True
    End If
End Sub
'+----------------------------------------+
'+  H1Op_Proc   処理区分
'+----------------------------------------+
Private Sub H1Op_Proc_Click(Index As Integer)
    
    H1Op_Proc(0).Tag = Index
    For Each gObj In H1Op_Proc
        If gObj.Index = Index Then
            gObj.BackColor = mGotColor
        Else
            gObj.BackColor = mLostColor
        End If
    Next gObj
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

    RV_TorKb = 5    ' 得意先
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
       iH1_Torcd = RV_TorcdK
       Call iH1_Torcd_LostFocus
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
    If TorNmGet(iH1_Torcd, 2) Then
        H1lb_Torcd = TRM_RName
        If Trim(H1lb_Torcd) = "" Then
           H1lb_Torcd = TRM_Name
        End If
    Else
        H1lb_Torcd = ""
    End If
End Sub
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
            .MB_MSG(2) = "    今回対象年月を入力して下さい         "
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
'   < 締め日 >
    If H1Op_Shime(1).Value And CisFun.Val2(iH1_Shime) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    締日 個別の時、締日必須！         "
            .MB_Button = Error
            .MBOX
            iH1_Shime.SetFocus
            Exit Function
        End With
    End If
    
    With CisFun
        .MB_Lines = 4
        gStr = Left(iH1_YM, 4) & " 年 " & Mid(iH1_YM, 5, 2) & "月度 "
        If RTrim$(iH1_YM) = RTrim$(mZenYm) Then
            .MB_MSG(1) = "   " & gStr & " 【再処理】         "
        Else
            .MB_MSG(1) = "   " & gStr & " 【初回処理】         "
        End If
        .MB_MSG(3) = "    買掛実績作成処理を開始します？         "
        .MB_Button = Yes_No
        If Not .MBOX Then Exit Function
    End With
    
    Head1Chk = True
End Function
'*--------------------------*
'*      処　理　開　始      *
'*--------------------------*
Private Function Proc_Start() As Boolean
    Proc_Start = False
    
    STime = Now()
    mShime = CisFun.Val2(iH1_Shime)         ' 0:全て、99:末日
    mTorcd = RTrim(iH1_Torcd)
    ' 処理区分
    If H1Op_Proc(0).Value = True Then
       mSyori = ""                          ' 新規
    Else
       mSyori = "1"                         ' 再処理
    End If
    
    On Error GoTo Proc_Start_Err
'--( ワーク削除 )--
    CisDB.SQL = "TRUNCATE TABLE 買掛契約先ワーク"
    CisDB.DBExec
    
'======================================='
'   【購入】
'======================================='
    lb_msg = "【 購入実績処理中 】しばらくお待ち下さい": DoEvents
    DoEvents
    
    If Not KonyuProc Then Exit Function

'======================================='
'   【支給】
'======================================='
    lb_msg = "【 支給実績処理中 】しばらくお待ち下さい": DoEvents
    DoEvents
    
    If Not ShikyuProc Then Exit Function

'========================================== ***2008/06/13 Start
'======================================='
'   【材料購入】
'======================================='
    lb_msg = "【 材料購入実績処理中 】しばらくお待ち下さい": DoEvents
    DoEvents
    
    If Not ZKonyuProc Then Exit Function

'======================================='
'   【材料支給】
'======================================='
    lb_msg = "【 材料支給実績処理中 】しばらくお待ち下さい": DoEvents
    DoEvents
    
    If Not ZShikyuProc Then Exit Function

'========================================== ***2008/06/13 End

'======================================='
'   【相殺処理】　支払テーブル作成
'======================================='
    lb_msg = "【 相殺処理中 】しばらくお待ち下さい": DoEvents
    DoEvents

    If Not SousaiProc Then Exit Function

Proc_Start_ED:

'   < 今回内容保存 >
    Call ProcRegistry_Save
    
    On Error GoTo 0
    
    Proc_Start = True
    Exit Function

Proc_Start_Err:
    CisFun.ErrorBox "買掛実績作成処理"
    End
End Function
'/////////////////////////////////////////////////////
'+---------------------------------------------------+
'+          購入実績処理
'+---------------------------------------------------+
'/////////////////////////////////////////////////////
Private Function KonyuProc() As Boolean
    KonyuProc = False
    
'''    Call CisDB.DBTran(TransBegin)
    With CisDB
        .SQL = "買掛購入実績作成"
        .StoadoCount = 14
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用
        .ParaNo = 5:    .ParaIO = Input用
        .ParaNo = 6:    .ParaIO = OutPut用
        .ParaNo = 7:    .ParaIO = OutPut用
        .ParaNo = 8:    .ParaIO = OutPut用
        .ParaNo = 9:    .ParaIO = OutPut用
        .ParaNo = 10:   .ParaIO = OutPut用
        .ParaNo = 11:   .ParaIO = OutPut用
        .ParaNo = 12:   .ParaIO = OutPut用
        .ParaNo = 13:   .ParaIO = OutPut用
        .ParaNo = 14:   .ParaIO = OutPut用

        .ParaNo = 1:    .ParaValue = iH1_YM                         ' 対象年月
        .ParaNo = 2:    .ParaValue = mTorcd                         ' 契約先
        .ParaNo = 3:    .ParaValue = mShime                         ' 締日
        .ParaNo = 4:    .ParaValue = mSyori                         ' 処理("":新規、"1":再処理)
        .ParaNo = 5:    .ParaValue = gTanto                         ' 処理担当者

        .DBStored

        .ParaNo = 0
        If .ParaValue <> 0 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    購入実績作成処理　異常！      "
            CisFun.MB_MSG(4) = "    Error No = " & .ParaValue
            CisFun.MB_Button = Error
            CisFun.MBOX
            End
        End If

        .ParaNo = 7: C1lb_Input(1) = Format(CisFun.Val2(C1lb_Input(1)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 8: C1lb_Input(2) = Format(CisFun.Val2(C1lb_Input(2)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 9: C1lb_Input(1) = Format(CisFun.Val2(C1lb_Input(1)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 10: C1lb_Input(3) = Format(CisFun.Val2(C1lb_Input(3)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 11: C1lb_Output(2) = Format(CisFun.Val2(C1lb_Output(2)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 12: C1lb_Output(1) = Format(CisFun.Val2(C1lb_Output(1)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 14: C1lb_Input(5) = Format(CisFun.Val2(.ParaValue), "#,###")
    End With
'''    Call CisDB.DBTran(TransCommit)
    
    
    KonyuProc = True
End Function
'/////////////////////////////////////////////////////
'+---------------------------------------------------+
'+          支給実績処理
'+---------------------------------------------------+
'/////////////////////////////////////////////////////
Private Function ShikyuProc() As Boolean
    ShikyuProc = False
    
'''    Call CisDB.DBTran(TransBegin)
    With CisDB
        .SQL = "買掛支給実績作成"
        .StoadoCount = 12
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用
        .ParaNo = 5:    .ParaIO = Input用
        .ParaNo = 6:    .ParaIO = OutPut用
        .ParaNo = 7:    .ParaIO = OutPut用
        .ParaNo = 8:    .ParaIO = OutPut用
        .ParaNo = 9:    .ParaIO = OutPut用
        .ParaNo = 10:   .ParaIO = OutPut用
        .ParaNo = 11:   .ParaIO = OutPut用
        .ParaNo = 12:   .ParaIO = OutPut用

        .ParaNo = 1:    .ParaValue = iH1_YM                         ' 対象年月
        .ParaNo = 2:    .ParaValue = mTorcd                         ' 契約先
        .ParaNo = 3:    .ParaValue = mShime                         ' 締日                       ' 契約先
        .ParaNo = 4:    .ParaValue = mSyori                         ' 処理("":新規、"1":再処理)
        .ParaNo = 5:    .ParaValue = gTanto                         ' 処理担当者

        .DBStored

        .ParaNo = 0
        If .ParaValue <> 0 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    支給実績作成処理　異常！      "
            CisFun.MB_MSG(4) = "    Error No = " & .ParaValue
            CisFun.MB_Button = Error
            CisFun.MBOX
            End
        End If

        .ParaNo = 7: C1lb_Input(0) = Format(CisFun.Val2(C1lb_Input(0)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 8: C1lb_Input(0) = Format(CisFun.Val2(C1lb_Input(0)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 9: C1lb_Input(4) = Format(CisFun.Val2(C1lb_Input(4)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 10: C1lb_Output(4) = Format(CisFun.Val2(C1lb_Output(4)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 11: C1lb_Output(0) = Format(CisFun.Val2(C1lb_Output(0)) + CisFun.Val2(.ParaValue), "#,###")
    End With
'''    Call CisDB.DBTran(TransCommit)
    
    
    ShikyuProc = True
End Function
'================================================================================ *** 2008/06/13 Start
'/////////////////////////////////////////////////////
'+---------------------------------------------------+
'+          材料購入実績処理
'+---------------------------------------------------+
'/////////////////////////////////////////////////////
Private Function ZKonyuProc() As Boolean
    ZKonyuProc = False
    
    With CisDB
        .SQL = "材料買掛購入実績作成"
        .StoadoCount = 12
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用
        .ParaNo = 5:    .ParaIO = Input用
        .ParaNo = 6:    .ParaIO = OutPut用
        .ParaNo = 7:    .ParaIO = OutPut用
        .ParaNo = 8:    .ParaIO = OutPut用
        .ParaNo = 9:    .ParaIO = OutPut用
        .ParaNo = 10:   .ParaIO = OutPut用
        .ParaNo = 11:   .ParaIO = OutPut用
        .ParaNo = 12:   .ParaIO = OutPut用

        .ParaNo = 1:    .ParaValue = iH1_YM                         ' 対象年月
        .ParaNo = 2:    .ParaValue = mTorcd                         ' 契約先
        .ParaNo = 3:    .ParaValue = mShime                         ' 締日
        .ParaNo = 4:    .ParaValue = mSyori                         ' 処理("":新規、"1":再処理)
        .ParaNo = 5:    .ParaValue = gTanto                         ' 処理担当者

        .DBStored

        .ParaNo = 0
        If .ParaValue <> 0 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    材料購入実績作成処理　異常！      "
            CisFun.MB_MSG(4) = "    Error No = " & .ParaValue
            CisFun.MB_Button = Error
            CisFun.MBOX
            End
        End If

        .ParaNo = 7: C1lb_Input(6) = Format(CisFun.Val2(C1lb_Input(6)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 8: C1lb_Input(6) = Format(CisFun.Val2(C1lb_Input(6)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 11: C1lb_Output(6) = Format(CisFun.Val2(C1lb_Output(6)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 10: C1lb_Output(5) = Format(CisFun.Val2(C1lb_Output(5)) + CisFun.Val2(.ParaValue), "#,###")
    End With
    
    ZKonyuProc = True
End Function
'/////////////////////////////////////////////////////
'+---------------------------------------------------+
'+          材料支給実績処理
'+---------------------------------------------------+
'/////////////////////////////////////////////////////
Private Function ZShikyuProc() As Boolean
    ZShikyuProc = False
    
    With CisDB
        .SQL = "材料買掛支給実績作成"
        .StoadoCount = 12
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用
        .ParaNo = 5:    .ParaIO = Input用
        .ParaNo = 6:    .ParaIO = OutPut用
        .ParaNo = 7:    .ParaIO = OutPut用
        .ParaNo = 8:    .ParaIO = OutPut用
        .ParaNo = 9:    .ParaIO = OutPut用
        .ParaNo = 10:   .ParaIO = OutPut用
        .ParaNo = 11:   .ParaIO = OutPut用
        .ParaNo = 12:   .ParaIO = OutPut用

        .ParaNo = 1:    .ParaValue = iH1_YM                         ' 対象年月
        .ParaNo = 2:    .ParaValue = mTorcd                         ' 契約先
        .ParaNo = 3:    .ParaValue = mShime                         ' 締日                       ' 契約先
        .ParaNo = 4:    .ParaValue = mSyori                         ' 処理("":新規、"1":再処理)
        .ParaNo = 5:    .ParaValue = gTanto                         ' 処理担当者

        .DBStored

        .ParaNo = 0
        If .ParaValue <> 0 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    材料支給実績作成処理　異常！      "
            CisFun.MB_MSG(4) = "    Error No = " & .ParaValue
            CisFun.MB_Button = Error
            CisFun.MBOX
            End
        End If

        .ParaNo = 7: C1lb_Input(7) = Format(CisFun.Val2(C1lb_Input(7)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 8: C1lb_Input(7) = Format(CisFun.Val2(C1lb_Input(7)) + CisFun.Val2(.ParaValue), "#,###")
        '.ParaNo = 9: C1lb_Input(4) = Format(CisFun.Val2(C1lb_Input(4)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 10: C1lb_Output(8) = Format(CisFun.Val2(C1lb_Output(8)) + CisFun.Val2(.ParaValue), "#,###")
        .ParaNo = 11: C1lb_Output(7) = Format(CisFun.Val2(C1lb_Output(7)) + CisFun.Val2(.ParaValue), "#,###")
    End With
    
    ZShikyuProc = True
End Function
'================================================================================ *** 2008/06/13 End
'/////////////////////////////////////////////////////
'+---------------------------------------------------+
'+          相殺処理
'+---------------------------------------------------+
'/////////////////////////////////////////////////////
Private Function SousaiProc() As Boolean
    SousaiProc = False
    
'''    Call CisDB.DBTran(TransBegin)
    With CisDB
        .SQL = "買掛相殺処理"
        .StoadoCount = 5
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = OutPut用
        .ParaNo = 4:    .ParaIO = OutPut用
        .ParaNo = 5:    .ParaIO = OutPut用

        .ParaNo = 1:    .ParaValue = iH1_YM                         ' 対象年月
        .ParaNo = 2:    .ParaValue = gTanto                         ' 処理担当者

        .DBStored

        .ParaNo = 0
        If .ParaValue <> 0 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    相殺処理　異常！      "
            CisFun.MB_MSG(4) = "    Error No = " & .ParaValue
            CisFun.MB_Button = Error
            CisFun.MBOX
            End
        End If

        .ParaNo = 5: C1lb_Output(3) = Format(CisFun.Val2(.ParaValue), "#,###")
    End With
'''    Call CisDB.DBTran(TransCommit)
    
    
    SousaiProc = True
End Function


'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽
'▽                         名　称　マ　ス　タ　確　認  ( シ ス テ ム 共 通 定 数 等 )
'▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'*---------------------------------------------------------*
'*                  前　回　内　容　獲　得                  *
'*---------------------------------------------------------*
Private Sub ProcRegistry_Check()
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "処理内容"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "買掛実績作成処理"
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
                H1lb_ZYM = Format(.英数字1, "@@@@年 @@月度")
                mZenYm = RTrim(.英数字1)
            End If
            If RTrim$(.英数字2) <> "" Then
                H1lb_ZDate = .英数字2
            End If
            If RTrim$(.英数字3) <> "" Then
                If TorNmGet(.英数字3, 2) Then
                    H1lb_KTorcd = .英数字3 & ":" & TRM_RName
                    If Trim(H1lb_KTorcd) = "" Then
                       H1lb_KTorcd = .英数字3 & ":" & TRM_Name
                    End If
                Else
                    H1lb_KTorcd = .英数字3
                End If
            End If
            Select Case .数字1
                Case 0: H1lb_Shime = "全て"
                Case 1: H1lb_Shime = "個別  " & Format(.数字2, "#")
                Case Else
                        H1lb_Shime = "個別(実績済みを除く)  " & Format(.数字2, "#")
            End Select
            If .数字3 = 0 Then
                H1lb_Proc = "新規"
            Else
                H1lb_Proc = "再処理"
            End If
        End With
    End If
End Sub
'*---------------------------------------------------------*
'*                  今　回　内　容　保  存                  *
'*---------------------------------------------------------*
Private Sub ProcRegistry_Save()

    Call ItemsClearSYM
    
    With SYM
        .区分名称 = "処理内容"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "買掛実績作成処理"
        .英数字1 = iH1_YM
        .英数字2 = Format(Now(), "YYYY年 MM月 DD日 ") & Format(Now(), "HH:NN")
        .英数字3 = RTrim(iH1_Torcd)
        .数字1 = CisFun.Val2(H1Op_Shime(0).Tag)
        .数字2 = CisFun.Val2(iH1_Shime)
        .数字3 = CisFun.Val2(H1Op_Proc(0).Tag)
        .作成者 = gTanto
    End With

    Call SYMUpdate
End Sub
'*---------------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス チ ェ ッ ク             *
'*---------------------------------------------------------*
Private Function SingleProcess_Check()
    SingleProcess_Check = False
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "シングルプロセス"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "買掛実績作成処理"
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
'*          シ ン グ ル プ ロ セ ス 解 除           *
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
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()

' 消費税
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "消費税"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "消費税"
        .英数字1 = "数字1=消費税率"
        .英数字2 = "数字2=端数処理 → 0:切り捨て、"
        .英数字3 = "1:切上、2:四捨五入"
        .数字1 = 5
        .制御区分 = ""
        .作成者 = gTanto
    End With
    Call SYMInitCreate(False)
' 繰越管理
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "買掛管理"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "繰越管理"
        .英数字1 = "数字1=0:しない、=1:する"
        .制御区分 = "X"
        .作成者 = gTanto
    End With
    Call SYMInitCreate(False)
' 支払区分
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "買掛支払区分"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "現金"
        .英数字1 = "何月後"
        .英数字2 = "何日(末日=0)"
        .英数字3 = "率"
        .数字1 = 1
        .数字2 = 0
        .数字3 = 100
        .制御区分 = "X"
        .作成者 = gTanto
        .検索名 = "支払区分検索"
        .コード名 = "CD"
        .名称名 = "支払区分名"
    End With
    Call SYMInitCreate(True)
    With SYMCreate
        .区分名称 = "買掛支払区分"
        .区分桁数 = 1
        .値 = "2"
        .値名称 = "手形"
        .数字1 = 3
        .数字2 = 0
        .数字3 = 0
        .制御区分 = "X"
        .作成者 = gTanto
    End With
    Call SYMInitCreate(False)

' 識別子
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "買掛識別子"
        .区分桁数 = 2
        .値 = "01"
        .値名称 = "部品費"
        .英数字1 = "K"
        .数字1 = 0
        .制御区分 = "X"
        .作成者 = gTanto
        
        .検索名 = "識別子(買掛)　検索"
        .コード名 = "識別子"
        .名称名 = "識別子名称"
    End With
    Call SYMInitCreate(True)
    With SYMCreate
        .値 = "02"
        .値名称 = "型費"
        .英数字1 = "K"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "04"
        .値名称 = "その他部品費"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "05"
        .値名称 = "その他仕入"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "06"
        .値名称 = "金型"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "07"
        .値名称 = "金型修繕費"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "08"
        .値名称 = "治工具"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "10"
        .値名称 = "訂正(黒)"
        .英数字1 = "K"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "11"
        .値名称 = "訂正(赤)"
        .英数字1 = "K"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "12"
        .値名称 = "遡り訂正"
        .英数字1 = "K"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)

    With SYMCreate
        .値 = "21"
        .値名称 = "支給部品費"
        .英数字1 = "S"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "24"
        .値名称 = "その他部品費"
        .英数字1 = "S"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "25"
        .値名称 = "その他支給"
        .英数字1 = "S"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "30"
        .値名称 = "支給訂正(黒)"
        .英数字1 = "S"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "31"
        .値名称 = "支給訂正(赤)"
        .英数字1 = "S"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "32"
        .値名称 = "支給遡り訂正"
        .英数字1 = "S"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    
    '2008/01/22追加
    With SYMCreate
        .値 = "41"
        .値名称 = "受給部品費"
        .英数字1 = "J"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "42"
        .値名称 = "その他受給部品費"
        .英数字1 = "J"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "43"
        .値名称 = "受給訂正(黒)"
        .英数字1 = "J"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "44"
        .値名称 = "受給訂正(赤)"
        .英数字1 = "J"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)

    '======================================= *** 2008/06/13 Start
    With SYMCreate
        .値 = "51"
        .値名称 = "材料費"
        .英数字1 = "K"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "55"
        .値名称 = "その他仕入(材料)"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "71"
        .値名称 = "支給材料費"
        .英数字1 = "S"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "75"
        .値名称 = "その他支給(材料)"
        .英数字1 = "S"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    '======================================= *** 2008/06/13 End
    
End Sub


