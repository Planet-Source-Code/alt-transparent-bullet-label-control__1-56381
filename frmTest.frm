VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Transparent Bullet Label Control"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin BulletLabelControl.BulletLabel BulletLabel13 
      Height          =   2025
      Left            =   120
      Top             =   1950
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   3572
      BulletColor     =   0
      BulletOutlineWidth=   0
      Caption         =   $"frmTest.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4020
      Width           =   1020
   End
   Begin BulletLabelControl.BulletLabel BulletLabel11 
      Height          =   225
      Left            =   2145
      Top             =   1215
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   397
      BulletColor     =   32896
      BulletStyle     =   11
      Caption         =   "StarOutline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel9 
      Height          =   225
      Left            =   2145
      Top             =   630
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   397
      BulletColor     =   16512
      BulletStyle     =   9
      Caption         =   "DiamondOutline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel8 
      Height          =   225
      Left            =   2145
      Top             =   345
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   397
      BulletColor     =   32768
      BulletStyle     =   8
      Caption         =   "DownArrow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel7 
      Height          =   225
      Left            =   2145
      Top             =   45
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   397
      BulletColor     =   32768
      BulletStyle     =   7
      Caption         =   "UpArrow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel6 
      Height          =   270
      Left            =   90
      Top             =   1515
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   476
      BulletColor     =   255
      BulletStyle     =   6
      Caption         =   "RightArrow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel5 
      Height          =   225
      Left            =   90
      Top             =   1215
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   397
      BulletColor     =   255
      BulletStyle     =   5
      Caption         =   "LeftArrow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel4 
      Height          =   225
      Left            =   90
      Top             =   930
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   397
      BulletColor     =   16711680
      BulletOutlineWidth=   2
      BulletStyle     =   4
      Caption         =   "SquareFilled"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel3 
      Height          =   300
      Left            =   90
      Top             =   630
      Width           =   1695
      _ExtentX        =   6059
      _ExtentY        =   3122
      BulletColor     =   16711680
      BulletOutlineWidth=   2
      BulletStyle     =   3
      Caption         =   "SquareOutline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel2 
      Height          =   225
      Left            =   90
      Top             =   360
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   397
      BulletStyle     =   2
      Caption         =   "DiscFilled"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel1 
      Height          =   225
      Left            =   90
      Top             =   45
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   397
      BulletStyle     =   1
      Caption         =   "DiscOutline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel10 
      Height          =   225
      Left            =   2145
      Top             =   930
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   397
      BulletColor     =   33023
      BulletStyle     =   10
      Caption         =   "DiamondFilled"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BulletLabelControl.BulletLabel BulletLabel12 
      Height          =   225
      Left            =   2145
      Top             =   1515
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   397
      BulletColor     =   8454143
      BulletStyle     =   12
      Caption         =   "StarFilled"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
     Unload Me
End Sub
