VERSION 5.00
Begin VB.Form FrmGreenEffect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Green Effect"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "FRMTILE1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSupportRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "FRMTILE1.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   59
      Top             =   8640
      Width           =   495
   End
   Begin VB.PictureBox PicSupportLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":0860
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   58
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicBed 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":0C85
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   57
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicStepsRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":10EE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   56
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicSteps 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":155A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   55
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicStepsLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":18E0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   54
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2280
      Picture         =   "FRMTILE1.frx":1D4B
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   53
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicCarpetRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2820
      Picture         =   "FRMTILE1.frx":21A9
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   52
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicBottomRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1740
      Picture         =   "FRMTILE1.frx":25FE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   51
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetBottom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   660
      Picture         =   "FRMTILE1.frx":2ABE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   50
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicCarpet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FRMTILE1.frx":2EEE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   49
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicBottomLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1200
      Picture         =   "FRMTILE1.frx":324B
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   48
      Top             =   8100
      Width           =   500
   End
   Begin VB.PictureBox PicCarpetTop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      Picture         =   "FRMTILE1.frx":370D
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   47
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3900
      Picture         =   "FRMTILE1.frx":3B45
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   46
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicCarpetTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4440
      Picture         =   "FRMTILE1.frx":3FFE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   45
      Top             =   8100
      Width           =   495
   End
   Begin VB.PictureBox PicArmor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":44B1
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   44
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicInn 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      Picture         =   "FRMTILE1.frx":48A8
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   43
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicCase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":4C7C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   42
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicBCase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":5085
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   41
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicSign 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      Picture         =   "FRMTILE1.frx":54EE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   40
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicPerson2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4380
      Picture         =   "FRMTILE1.frx":5900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   39
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicPerson1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4380
      Picture         =   "FRMTILE1.frx":5D56
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   38
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4440
      Picture         =   "FRMTILE1.frx":61A7
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   37
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicStool 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4380
      Picture         =   "FRMTILE1.frx":65FA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   36
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicDoor 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3900
      Picture         =   "FRMTILE1.frx":6A60
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   35
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicBrick 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3900
      Picture         =   "FRMTILE1.frx":6E47
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   34
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicDirt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      Picture         =   "FRMTILE1.frx":7214
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   33
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicDown 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   3840
      Picture         =   "FRMTILE1.frx":7641
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   32
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicDown 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   3840
      Picture         =   "FRMTILE1.frx":7AAB
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   31
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   3300
      Picture         =   "FRMTILE1.frx":7F12
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   30
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   3300
      Picture         =   "FRMTILE1.frx":8365
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   29
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   2760
      Picture         =   "FRMTILE1.frx":87B8
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   28
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   2760
      Picture         =   "FRMTILE1.frx":8BEF
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   27
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   2220
      Picture         =   "FRMTILE1.frx":903C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   26
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   2220
      Picture         =   "FRMTILE1.frx":9488
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   25
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicFlowers 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   3360
      Picture         =   "FRMTILE1.frx":98C4
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   7020
      Width           =   500
   End
   Begin VB.Timer TmrPlayer 
      Interval        =   100
      Left            =   5520
      Top             =   0
   End
   Begin VB.PictureBox PicBotLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FRMTILE1.frx":9DEB
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicStopleft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   660
      Picture         =   "FRMTILE1.frx":A31F
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   20
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicStopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1200
      Picture         =   "FRMTILE1.frx":A849
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   19
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicBotRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FRMTILE1.frx":AD6F
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   18
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicFenceLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   660
      Picture         =   "FRMTILE1.frx":B2A8
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   17
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicAcross 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1200
      Picture         =   "FRMTILE1.frx":B7E9
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   16
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicFenceRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1740
      Picture         =   "FRMTILE1.frx":BD1A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   7020
      Width           =   495
   End
   Begin VB.PictureBox PicStopLeftUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1740
      Picture         =   "FRMTILE1.frx":C255
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   7560
      Width           =   495
   End
   Begin VB.PictureBox PicChest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2820
      Picture         =   "FRMTILE1.frx":C794
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicStopRightUp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2280
      Picture         =   "FRMTILE1.frx":CC8C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   7560
      Width           =   500
   End
   Begin VB.PictureBox PicTopRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2820
      Picture         =   "FRMTILE1.frx":D1C8
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicTopLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   2280
      Picture         =   "FRMTILE1.frx":D6F5
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   7020
      Width           =   500
   End
   Begin VB.PictureBox PicWell 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1740
      Picture         =   "FRMTILE1.frx":DC1F
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   6480
      Width           =   495
   End
   Begin VB.PictureBox PicTree 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1740
      Picture         =   "FRMTILE1.frx":E162
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   5940
      Width           =   495
   End
   Begin VB.PictureBox PicWater 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1200
      Picture         =   "FRMTILE1.frx":E6B1
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   5940
      Width           =   500
   End
   Begin VB.PictureBox PicRock 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   660
      Picture         =   "FRMTILE1.frx":15739
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   6480
      Width           =   500
   End
   Begin VB.PictureBox PicSand 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FRMTILE1.frx":19B7D
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   6480
      Width           =   500
   End
   Begin VB.PictureBox PicWater1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   1200
      Picture         =   "FRMTILE1.frx":1A083
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   6480
      Width           =   500
   End
   Begin VB.PictureBox PicPath 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   660
      Picture         =   "FRMTILE1.frx":2A4A7
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   5940
      Width           =   500
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   120
      Picture         =   "FRMTILE1.frx":2A8E1
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   5940
      Width           =   500
   End
   Begin VB.Image Imgplayer 
      Height          =   465
      Left            =   0
      Picture         =   "FRMTILE1.frx":2AE0D
      Top             =   0
      Width           =   465
   End
   Begin VB.Label LblAble 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   4260
      TabIndex        =   24
      Top             =   3420
      Width           =   1575
   End
   Begin VB.Label LblPos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   4260
      TabIndex        =   23
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label LblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This is an rpg created by kevin pfister in 2002, this dialog is the test section for the main app. with help from ian fulford*"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   5595
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   3855
      Left            =   4140
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   1635
      Left            =   120
      Top             =   4140
      Width           =   5835
   End
   Begin VB.Label LblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Green Efffect"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4260
      Width           =   5595
   End
End
Attribute VB_Name = "FrmGreenEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalMap(1 To 200) As String    'The Map array
Dim CompressedMap(1 To 200) As String    'The Map array
Dim Map(1 To 10) As String          'The Current Map onScreen

Dim MapPlayers(1 To 200, 1 To 200) As Integer
Dim PlayerName(1 To 800) As String
Dim PlayerText(1 To 800) As String
Dim PlayerType(1 To 800) As Integer
Dim CharX(1 To 800) As Integer
Dim CharY(1 To 800) As Integer
Dim HouseName(1 To 800) As String
Dim HouseX(1 To 800) As Integer
Dim HouseY(1 To 800) As Integer
Dim MapHouse(1 To 200, 1 To 200) As Integer

Dim MapGrid(1 To 10, 1 To 10) As String
Dim Color, Color2
Dim MapText(1 To 200) As String
Dim r As Integer, g As Integer, b As Integer
Dim rback As Integer, gback As Integer, bback As Integer
Dim ScreenX
Dim ScreenY
Dim PlayerX
Dim PlayerY
Dim MessageValue
Dim Anim
Dim InBuilding As Boolean
Dim BuildingX
Dim BuildingY
Dim File As String
Dim DefFolder As String

Dim NiceGradient As Boolean

Private Sub Form_Load()
    Randomize
    PlayerX = 4
    PlayerY = 4
    Imgplayer.Top = PlayerY * 375 - 240
    Imgplayer.Left = PlayerX * 375 - 135
    ScreenX = 1
    ScreenY = 1
    Anim = 1
    InBuilding = False
    
    'This must be right in order for the game to work, it is the Game Data Folder
    
    DefFolder = "C:\Windows\Desktop\Game Data\"    '<---------- The Game Directory
    
    
    
    NiceGradient = True 'False = Faster less graphics... True = Slower Better Grpahics
    
    'A = Wall
    'B = Door
    'C = Chest
    'D = Dirt
    'E = Grass Rocks
    'F = Flowers
    'G = Grass
    'H = Stool
    'I = Window
    'J = Person1
    'K = Person2
    'L = Book Case
    'M = Case
    'N = Carpet
    'O = Carpet Top Left
    'P = Path
    'Q = Carpet Top
    'R = Rock
    'S = Sand
    'T = Tree
    'U = Carpet Top Right
    'V = Carpet Right
    'W = Water
    'X = Carpet Bottom
    'Y = Carpet Bottom Left
    'Z = Carpet Left
    '1 = Bottom Left Fence
    '2 = Bottom Right Fence
    '3 = Stop at left Fence
    '4 = Left Fence
    '5 = Horizontal Fence
    '6 = Stop at right fence
    '7 = Right Fence
    '8 = Stop at Upper Left
    '9 = Top Left Fence
    '0 = Stop at Upper Right
    '/ = Top Right Fence
    '! = Steps Left
    '% = Steps
    '£ = Steps Right
    '$ = Bed
    '^ = Carpet Bottom Right
    '& = Support Left
    '* = Support Right
    
    Open DefFolder + "GEffect.Map" For Input As #1
    ' If the Path is not right change the DefFolder Path
    
    For OuterLoop = 1 To 200
        Input #1, CompressedMap(OuterLoop)
    Next
    Close
    
    For OuterLoop = 1 To 200
        TotalMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            If Mid(CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    Open DefFolder + "GEffect.Pls" For Input As #1
    ' If the Path is not right change the DefFolder Path
    CurrentPlayer = 0
    Input #1, Text
    ToNum = Val(Text)
    For Extras = 1 To ToNum
        For GetInfo = 1 To 5
            Input #1, Text
            If Mid(Text, 1, 6) = "##Name" Then
                CurrentPlayer = CurrentPlayer + 1
                PlayerName(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##XPos" Then
                CharX(CurrentPlayer) = Val(Mid(Text, 8))
            ElseIf Mid(Text, 1, 6) = "##YPos" Then
                CharY(CurrentPlayer) = Val(Mid$(Text, 8))
                MapPlayers(CharX(CurrentPlayer), CharY(CurrentPlayer)) = CurrentPlayer
            ElseIf Mid(Text, 1, 6) = "##Text" Then
                PlayerText(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##Type" Then
                PlayerType(CurrentPlayer) = Val(Mid(Text, 8))
            End If
        Next
    Next
    Close

    Open DefFolder + "GEffect.Hus" For Input As #1
    ' If the Path is not right change the DefFolder Path
    CurrentHouse = 0
    Input #1, Text
    ToNum = Val(Text)
    For Extras = 1 To ToNum
        For GetInfo = 1 To 3
            Input #1, Text
            If Mid(Text, 1, 6) = "##Name" Then
                CurrentHouse = CurrentHouse + 1
                HouseName(CurrentHouse) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##XPos" Then
                HouseX(CurrentHouse) = Val(Mid(Text, 8))
            ElseIf Mid(Text, 1, 6) = "##YPos" Then
                HouseY(CurrentHouse) = Val(Mid$(Text, 8))
                MapHouse(HouseX(CurrentHouse), HouseY(CurrentHouse)) = CurrentHouse
            End If
        Next
    Next
    Close

    FrmGreenEffect.Show
    FrmGreenEffect.Visible = True
    RenderMap
End Sub


Sub RenderMap()
    TmrPlayer.Enabled = False
    For OuterLoop = 1 To 10
        Map(OuterLoop) = ""
        For InnerLoop = 1 To 10
            MapGrid(InnerLoop, OuterLoop) = ""
        Next
    Next
    
    For OuterLoop = 1 To 10
        For InnerLoop = 1 To 10
            Map(OuterLoop) = Map(OuterLoop) + Mid$(TotalMap(ScreenY * 10 - 10 + OuterLoop), ScreenX * 10 - 10 + InnerLoop, 1)
            MapGrid(InnerLoop, OuterLoop) = Mid$(Map(OuterLoop), InnerLoop, 1)
            If MapGrid(InnerLoop, OuterLoop) = "G" Then
                Call FrmGreenEffect.PaintPicture(PicGrass, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "W" Then
                Call FrmGreenEffect.PaintPicture(PicWater, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "S" Then
                Call FrmGreenEffect.PaintPicture(PicSand, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "R" Then
                Call FrmGreenEffect.PaintPicture(PicRock, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "T" Then
                Call FrmGreenEffect.PaintPicture(PicTree, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "E" Then
                Call FrmGreenEffect.PaintPicture(PicWell, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "C" Then
                Call FrmGreenEffect.PaintPicture(PicChest, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "P" Then
                Call FrmGreenEffect.PaintPicture(PicPath, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "F" Then
                Call FrmGreenEffect.PaintPicture(PicFlowers, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "D" Then
                Call FrmGreenEffect.PaintPicture(PicDirt, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "A" Then
                Call FrmGreenEffect.PaintPicture(PicBrick, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "B" Then
                Call FrmGreenEffect.PaintPicture(PicDoor, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "H" Then
                Call FrmGreenEffect.PaintPicture(PicStool, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "I" Then
                Call FrmGreenEffect.PaintPicture(PicWindow, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "1" Then
                Call FrmGreenEffect.PaintPicture(PicBotLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "2" Then
                Call FrmGreenEffect.PaintPicture(PicBotRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "3" Then
                Call FrmGreenEffect.PaintPicture(PicStopleft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "4" Then
                Call FrmGreenEffect.PaintPicture(PicFenceLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "5" Then
                Call FrmGreenEffect.PaintPicture(PicAcross, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "6" Then
                Call FrmGreenEffect.PaintPicture(PicStopRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "7" Then
                Call FrmGreenEffect.PaintPicture(PicFenceRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "8" Then
                Call FrmGreenEffect.PaintPicture(PicStopLeftUp, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "9" Then
                Call FrmGreenEffect.PaintPicture(PicTopLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "0" Then
                Call FrmGreenEffect.PaintPicture(PicStopRightUp, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "/" Then
                Call FrmGreenEffect.PaintPicture(PicTopRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "L" Then
                Call FrmGreenEffect.PaintPicture(PicBCase, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "M" Then
                Call FrmGreenEffect.PaintPicture(PicCase, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "N" Then
                Call FrmGreenEffect.PaintPicture(PicCarpet, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "O" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetTopLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "U" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetTopRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "V" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "^" Then
                Call FrmGreenEffect.PaintPicture(PicBottomRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "X" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetBottom, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "Y" Then
                Call FrmGreenEffect.PaintPicture(PicBottomLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "Z" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "Q" Then
                Call FrmGreenEffect.PaintPicture(PicCarpetTop, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "!" Then
                Call FrmGreenEffect.PaintPicture(PicStepsLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "%" Then
                Call FrmGreenEffect.PaintPicture(PicSteps, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "£" Then
                Call FrmGreenEffect.PaintPicture(PicStepsRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "$" Then
                Call FrmGreenEffect.PaintPicture(PicBed, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "&" Then
                Call FrmGreenEffect.PaintPicture(PicSupportLeft, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            ElseIf MapGrid(InnerLoop, OuterLoop) = "*" Then
                Call FrmGreenEffect.PaintPicture(PicSupportRight, InnerLoop * 375 - 135, OuterLoop * 375 - 135, 400, 400)
            End If
        Next
    Next
    If InBuilding = False And NiceGradient = True Then
        For OutFade = 1 To 10
            For Infade = 1 To 10
                If Infade - 1 > 0 Then
                    If MapGrid(Infade, OutFade) <> MapGrid(Infade - 1, OutFade) Then
                        For i = 1 To 50 Step 15
                            For ii = 1 To 375 Step 15
                                If MapGrid(Infade - 1, OutFade) = "G" Then
                                    Color = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "W" Then
                                    Color = PicWater.Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "S" Then
                                    Color = PicSand.Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "R" Then
                                    Color = PicRock.Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "T" Then
                                    Color = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "E" Then
                                    Color = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "P" Then
                                    Color = PicPath.Point(i, ii)
                                ElseIf MapGrid(Infade - 1, OutFade) = "D" Then
                                    Color = PicDirt.Point(i, ii)
                                Else
                                    Color = 0
                                End If
                                If MapGrid(Infade, OutFade) = "G" Then
                                    Color2 = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "W" Then
                                    Color2 = PicWater.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "S" Then
                                    Color2 = PicSand.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "R" Then
                                    Color2 = PicRock.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "T" Then
                                    Color2 = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "E" Then
                                    Color2 = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "P" Then
                                    Color2 = PicPath.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "D" Then
                                    Color2 = PicDirt.Point(i, ii)
                                Else
                                    Color2 = 0
                                End If
                                If Color <> 0 And Color2 <> 0 Then
                                    GetRgb Color, r, g, b
                                    GetRgb Color2, rback, gback, bback
                                    Percent = 100 / ((i / 50) * 100)
                                    FrmGreenEffect.PSet (375 * Infade - 135 - i, 375 * OutFade + 240 - ii), RGB(rback - rback / Percent + r / Percent, gback - gback / Percent + g / Percent, bback - bback / Percent + b / Percent)
                                End If
                            Next
                        Next
                    End If
                End If
                If OutFade - 1 > 0 Then
                    If MapGrid(Infade, OutFade) <> MapGrid(Infade, OutFade - 1) Then
                        For i = 1 To 375 Step 15
                            For ii = 1 To 50 Step 15
                                If MapGrid(Infade, OutFade - 1) = "G" Then
                                    Color = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "W" Then
                                    Color = PicWater.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "S" Then
                                    Color = PicSand.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "R" Then
                                    Color = PicRock.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "T" Then
                                    Color = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "E" Then
                                    Color = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "P" Then
                                    Color = PicPath.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade - 1) = "D" Then
                                    Color = PicDirt.Point(i, ii)
                                Else
                                    Color = 0
                                End If
                                If MapGrid(Infade, OutFade) = "G" Then
                                    Color2 = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "W" Then
                                    Color2 = PicWater.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "S" Then
                                    Color2 = PicSand.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "R" Then
                                    Color2 = PicRock.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "T" Then
                                    Color2 = PicTree.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "E" Then
                                    Color2 = PicGrass.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "P" Then
                                    Color2 = PicPath.Point(i, ii)
                                ElseIf MapGrid(Infade, OutFade) = "D" Then
                                    Color2 = PicDirt.Point(i, ii)
                                Else
                                    Color2 = 0
                                End If
                                If Color <> 0 And Color2 <> 0 Then
                                    GetRgb Color, r, g, b
                                    GetRgb Color2, rback, gback, bback
                                    Percent = 100 / ((ii / 50) * 100)
                                    FrmGreenEffect.PSet (375 * Infade + 240 - i, 375 * OutFade - 135 - ii), RGB(rback - rback / Percent + r / Percent, gback - gback / Percent + g / Percent, bback - bback / Percent + b / Percent)
                                End If
                            Next
                        Next
                    End If
                End If
            Next
        Next
    End If
    
    For OuterLoop = 1 To 10
        For InnerLoop = 1 To 10
            If MapPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop)) > 0 Then
                    For InOuterLoop = 1 To 495 Step 15
                        For InInnerLoop = 1 To 495 Step 15
                            If PlayerType(MapPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 1 Then
                                Color = PicPerson1.Point(InInnerLoop, InOuterLoop)
                            ElseIf PlayerType(MapPlayers((ScreenX * 10 - 10 + InnerLoop), (ScreenY * 10 - 10 + OuterLoop))) = 2 Then
                                Color = PicPerson2.Point(InInnerLoop, InOuterLoop)
                            Else
                                Color = PicSign.Point(InInnerLoop, InOuterLoop)
                            End If
                            If Color <> vbGreen Then
                                FrmGreenEffect.PSet (InnerLoop * 375 - 135 + InInnerLoop, (OuterLoop - 1) * 375 - 135 + InOuterLoop), Color
                            End If
                        Next
                    Next
            End If
        Next
    Next
    TmrPlayer.Enabled = True
End Sub

Sub GetRgb(ByVal Color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
    Dim temp As Long
    temp = (Color And 255)
    red = temp And 255
    temp = Int(Color / 256)
    green = temp And 255
    temp = Int(Color / 65536)
    blue = temp And 255
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub TmrPlayer_Timer()
    If GetAsyncKeyState(vbKeyDown) < 0 Then
        Imgplayer.Picture = PicDown(Anim).Picture
        If InBuilding = True Then
            Call CheckBuilding
        End If
        Anim = 1 - Anim
        If PlayerY + 1 < 11 Then
            If MapGrid(PlayerX, PlayerY + 1) <> "R" And MapGrid(PlayerX, PlayerY + 1) <> "W" And MapGrid(PlayerX, PlayerY + 1) <> "E" And MapGrid(PlayerX, PlayerY + 1) <> "C" And MapGrid(PlayerX, PlayerY + 1) <> "T" And MapGrid(PlayerX, PlayerY + 1) <> "1" And MapGrid(PlayerX, PlayerY + 1) <> "2" And MapGrid(PlayerX, PlayerY + 1) <> "3" And MapGrid(PlayerX, PlayerY + 1) <> "4" And MapGrid(PlayerX, PlayerY + 1) <> "5" And MapGrid(PlayerX, PlayerY + 1) <> "6" And MapGrid(PlayerX, PlayerY + 1) <> "7" And MapGrid(PlayerX, PlayerY + 1) <> "8" And MapGrid(PlayerX, PlayerY + 1) <> "9" And MapGrid(PlayerX, PlayerY + 1) <> "0" And MapGrid(PlayerX, PlayerY + 1) <> "/" And MapGrid(PlayerX, PlayerY + 1) <> "A" And MapGrid(PlayerX, PlayerY + 1) <> "B" And MapGrid(PlayerX, PlayerY + 1) <> "H" And MapGrid(PlayerX, PlayerY + 1) <> "I" And MapGrid(PlayerX, PlayerY + 1) <> "L" And MapGrid(PlayerX, PlayerY + 1) <> "M" And MapGrid(PlayerX, PlayerY + 1) <> "$" And MapGrid(PlayerX, PlayerY + 1) <> "&" _
And MapGrid(PlayerX, PlayerY + 1) <> "*" Then
                PlayerY = PlayerY + 1
                Imgplayer.Top = PlayerY * 375 - 240
            End If
        Else
            PlayerY = 1
            ScreenY = ScreenY + 1
            RenderMap
            Imgplayer.Top = PlayerY * 375 - 240
        End If
    End If
    If GetAsyncKeyState(vbKeyUp) < 0 Then
        Imgplayer.Picture = PicUp(Anim).Picture
        If InBuilding = False Then
            Call CheckForBuilding
        End If
        Anim = 1 - Anim
        If PlayerY - 1 > 0 Then
            If MapGrid(PlayerX, PlayerY - 1) <> "R" And MapGrid(PlayerX, PlayerY - 1) <> "W" And MapGrid(PlayerX, PlayerY - 1) <> "E" And MapGrid(PlayerX, PlayerY - 1) <> "C" And MapGrid(PlayerX, PlayerY - 1) <> "T" And MapGrid(PlayerX, PlayerY - 1) <> "1" And MapGrid(PlayerX, PlayerY - 1) <> "2" And MapGrid(PlayerX, PlayerY - 1) <> "3" And MapGrid(PlayerX, PlayerY - 1) <> "4" And MapGrid(PlayerX, PlayerY - 1) <> "5" And MapGrid(PlayerX, PlayerY - 1) <> "6" And MapGrid(PlayerX, PlayerY - 1) <> "7" And MapGrid(PlayerX, PlayerY - 1) <> "8" And MapGrid(PlayerX, PlayerY - 1) <> "9" And MapGrid(PlayerX, PlayerY - 1) <> "0" And MapGrid(PlayerX, PlayerY - 1) <> "/" And MapGrid(PlayerX, PlayerY - 1) <> "A" And MapGrid(PlayerX, PlayerY - 1) <> "B" And MapGrid(PlayerX, PlayerY - 1) <> "H" And MapGrid(PlayerX, PlayerY - 1) <> "I" And MapGrid(PlayerX, PlayerY - 1) <> "L" And MapGrid(PlayerX, PlayerY - 1) <> "M" And MapGrid(PlayerX, PlayerY - 1) <> "$" And MapGrid(PlayerX, PlayerY - 1) <> "&" _
And MapGrid(PlayerX, PlayerY - 1) <> "*" Then
                PlayerY = PlayerY - 1
                Imgplayer.Top = PlayerY * 375 - 240
            End If
        Else
            PlayerY = 10
            ScreenY = ScreenY - 1
            RenderMap
            Imgplayer.Top = PlayerY * 375 - 240
        End If
    End If
    If GetAsyncKeyState(vbKeyLeft) < 0 Then
        Imgplayer.Picture = PicLeft(Anim).Picture
        Anim = 1 - Anim
        If PlayerX - 1 > 0 Then
            If MapGrid(PlayerX - 1, PlayerY) <> "R" And MapGrid(PlayerX - 1, PlayerY) <> "W" And MapGrid(PlayerX - 1, PlayerY) <> "E" And MapGrid(PlayerX - 1, PlayerY) <> "C" And MapGrid(PlayerX - 1, PlayerY) <> "T" And MapGrid(PlayerX - 1, PlayerY) <> "1" And MapGrid(PlayerX - 1, PlayerY) <> "2" And MapGrid(PlayerX - 1, PlayerY) <> "3" And MapGrid(PlayerX - 1, PlayerY) <> "4" And MapGrid(PlayerX - 1, PlayerY) <> "5" And MapGrid(PlayerX - 1, PlayerY) <> "6" And MapGrid(PlayerX - 1, PlayerY) <> "7" And MapGrid(PlayerX - 1, PlayerY) <> "8" And MapGrid(PlayerX - 1, PlayerY) <> "9" And MapGrid(PlayerX - 1, PlayerY) <> "0" And MapGrid(PlayerX - 1, PlayerY) <> "/" And MapGrid(PlayerX - 1, PlayerY) <> "A" And MapGrid(PlayerX - 1, PlayerY) <> "B" And MapGrid(PlayerX - 1, PlayerY) <> "H" And MapGrid(PlayerX - 1, PlayerY) <> "I" And MapGrid(PlayerX - 1, PlayerY) <> "L" And MapGrid(PlayerX - 1, PlayerY) <> "M" And MapGrid(PlayerX - 1, PlayerY) <> "$" And MapGrid(PlayerX - 1, PlayerY) <> "&" _
And MapGrid(PlayerX - 1, PlayerY) <> "*" Then
                PlayerX = PlayerX - 1
                Imgplayer.Left = PlayerX * 375 - 135
            End If
        Else
            PlayerX = 10
            ScreenX = ScreenX - 1
            RenderMap
            Imgplayer.Left = PlayerX * 375 - 135
        End If
    End If
    If GetAsyncKeyState(vbKeyRight) < 0 Then
    Imgplayer.Picture = PicRight(Anim).Picture
    Anim = 1 - Anim
        If PlayerX + 1 < 11 Then
            If MapGrid(PlayerX + 1, PlayerY) <> "R" And MapGrid(PlayerX + 1, PlayerY) <> "W" And MapGrid(PlayerX + 1, PlayerY) <> "E" And MapGrid(PlayerX + 1, PlayerY) <> "C" And MapGrid(PlayerX + 1, PlayerY) <> "T" And MapGrid(PlayerX + 1, PlayerY) <> "1" And MapGrid(PlayerX + 1, PlayerY) <> "2" And MapGrid(PlayerX + 1, PlayerY) <> "3" And MapGrid(PlayerX + 1, PlayerY) <> "4" And MapGrid(PlayerX + 1, PlayerY) <> "5" And MapGrid(PlayerX + 1, PlayerY) <> "6" And MapGrid(PlayerX + 1, PlayerY) <> "7" And MapGrid(PlayerX + 1, PlayerY) <> "8" And MapGrid(PlayerX + 1, PlayerY) <> "9" And MapGrid(PlayerX + 1, PlayerY) <> "0" And MapGrid(PlayerX + 1, PlayerY) <> "/" And MapGrid(PlayerX + 1, PlayerY) <> "A" And MapGrid(PlayerX + 1, PlayerY) <> "B" And MapGrid(PlayerX + 1, PlayerY) <> "H" And MapGrid(PlayerX + 1, PlayerY) <> "I" And MapGrid(PlayerX + 1, PlayerY) <> "L" And MapGrid(PlayerX + 1, PlayerY) <> "M" And MapGrid(PlayerX + 1, PlayerY) <> "$" And MapGrid(PlayerX + 1, PlayerY) <> "&" _
And MapGrid(PlayerX + 1, PlayerY) <> "*" Then
                PlayerX = PlayerX + 1
                Imgplayer.Left = PlayerX * 375 - 135
            End If
        Else
            PlayerX = 1
            ScreenX = ScreenX + 1
            RenderMap
            Imgplayer.Left = PlayerX * 375 - 135
        End If
    End If
    LblPos.Caption = Str$(ScreenX * 10 - 10 + PlayerX) + "," + Str$(ScreenY * 10 - 10 + PlayerY)
    CheckPos
End Sub

Sub CheckPos()
    LblAble.Caption = ""
    If MapPlayers((ScreenX * 10 - 10 + PlayerX), (ScreenY * 10 - 10 + PlayerY)) > 0 Then
        LblAble.Caption = "Message!"
        If GetAsyncKeyState(32) < 0 Then
            Call Chat((ScreenX * 10 - 10 + PlayerX), (ScreenY * 10 - 10 + PlayerY))
        End If
    End If
End Sub

Sub CheckBuilding()
    If ScreenX = 11 And ScreenY = 11 And PlayerX = 5 And PlayerY = 10 Then
        Call OutofBuilding
    End If
End Sub

Sub CheckForBuilding()
    If MapHouse((ScreenX * 10 - 10 + PlayerX), (ScreenY * 10 - 10 + PlayerY)) > 0 Then
        Call BuildingMaps(MapHouse((ScreenX * 10 - 10 + PlayerX), (ScreenY * 10 - 10 + PlayerY)))
    End If
End Sub

Sub BuildingMaps(Index)
    TmrPlayer.Enabled = False
    BuildingX = (ScreenX * 10 - 10 + PlayerX)
    BuildingY = (ScreenY * 10 - 11 + PlayerY)
    
    FileName = DefFolder + "GMap"
    
    File = FileName + Mid$(Str(Index), 2)
    ' If the Path is not right change the DefFolder Path
    Open File + ".Map" For Input As #1
    For OuterLoop = 1 To 200
        Input #1, CompressedMap(OuterLoop)
    Next
    Close
    
    ClearData
    
    For OuterLoop = 1 To 200
        TotalMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            If Mid(CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    
    Open File + ".Pls" For Input As #1
    CurrentPlayer = 0
    Input #1, Text
    ToNum = Val(Text)
    For Extras = 1 To ToNum
        For GetInfo = 1 To 5
            Input #1, Text
            If Mid(Text, 1, 6) = "##Name" Then
                CurrentPlayer = CurrentPlayer + 1
                PlayerName(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##XPos" Then
                CharX(CurrentPlayer) = Val(Mid(Text, 8))
            ElseIf Mid(Text, 1, 6) = "##YPos" Then
                CharY(CurrentPlayer) = Val(Mid$(Text, 8))
                MapPlayers(CharX(CurrentPlayer), CharY(CurrentPlayer)) = CurrentPlayer
            ElseIf Mid(Text, 1, 6) = "##Text" Then
                PlayerText(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##Type" Then
                PlayerType(CurrentPlayer) = Val(Mid(Text, 8))
            End If
        Next
    Next
    Close
    Imgplayer.Visible = False
    PlayerX = 5
    PlayerY = 11
    Imgplayer.Top = PlayerY * 375 - 240
    Imgplayer.Left = PlayerX * 375 - 135
    ScreenX = 11
    ScreenY = 11
    Anim = 1
    InBuilding = True
    RenderMap
    Imgplayer.Visible = True
    TmrPlayer.Enabled = True
End Sub

Sub Chat(InnerLoop, OuterLoop)
    Dim Chat(1 To 100) As String
    If PlayerName(MapPlayers(InnerLoop, OuterLoop)) = "SignPost" Then
        LblMessage.Caption = "Reading SignPost"
    Else
        LblMessage.Caption = "Talking to " + PlayerName(MapPlayers(InnerLoop, OuterLoop))
    End If
    If PlayerText(MapPlayers(InnerLoop, OuterLoop)) = "Random" Then
        Chat(1) = "Hello"
        Chat(2) = "Go Away"
        Chat(3) = "Welcome Traveller"
        Chat(4) = "Have you Voted yet?"
        Chat(5) = "Welcome to Muncipium"
        Chat(6) = "Have a look around"
        Chat(7) = "Have you saved lately"
        Chat(8) = "My Name is " + PlayerName(MapPlayers(InnerLoop, OuterLoop))
        Chat(9) = ""
        Chat(10) = ""
    
        RandomChat = Int(Rnd * 10) + 1
        LblText.Caption = Chat(RandomChat)
    Else
        LblText.Caption = PlayerText(MapPlayers(InnerLoop, OuterLoop))
    End If
    
    TimeBefore = Timer
    While Timer < TimeBefore + 2
        DoEvents
    Wend
    LblMessage.Caption = ""
    LblText.Caption = ""
End Sub

Sub OutofBuilding()
    TmrPlayer.Enabled = False
    InBuilding = False
    Open DefFolder + "GEffect.Map" For Input As #1
    ' If the Path is not right change the DefFolder Path
    For OuterLoop = 1 To 200
        Input #1, CompressedMap(OuterLoop)
    Next
    Close
    
    ClearData
    
    For OuterLoop = 1 To 200
        TotalMap(OuterLoop) = ""
        For InnerLoop = 1 To 200
            If Mid(CompressedMap(OuterLoop), InnerLoop, 1) = "(" Then
                If Mid(CompressedMap(OuterLoop), InnerLoop + 2, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 1))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1)
                    Next
                    InnerLoop = InnerLoop + 3
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 3, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 2))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1)
                    Next
                    InnerLoop = InnerLoop + 4
                ElseIf Mid(CompressedMap(OuterLoop), InnerLoop + 4, 1) = ")" Then
                    For A = 1 To Val(Mid(CompressedMap(OuterLoop), InnerLoop + 1, 3))
                        TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop + 5, 1)
                    Next
                    InnerLoop = InnerLoop + 5
                End If
            Else
                TotalMap(OuterLoop) = TotalMap(OuterLoop) + Mid(CompressedMap(OuterLoop), InnerLoop, 1)
            End If
        Next
    Next
    
    Open DefFolder + "GEffect.Pls" For Input As #1
    ' If the Path is not right change the DefFolder Path
    CurrentPlayer = 0
    Input #1, Text
    ToNum = Val(Text)
    For Extras = 1 To ToNum
        For GetInfo = 1 To 5
            Input #1, Text
            If Mid(Text, 1, 6) = "##Name" Then
                CurrentPlayer = CurrentPlayer + 1
                PlayerName(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##XPos" Then
                CharX(CurrentPlayer) = Val(Mid(Text, 8))
            ElseIf Mid(Text, 1, 6) = "##YPos" Then
                CharY(CurrentPlayer) = Val(Mid$(Text, 8))
                MapPlayers(CharX(CurrentPlayer), CharY(CurrentPlayer)) = CurrentPlayer
            ElseIf Mid(Text, 1, 6) = "##Text" Then
                PlayerText(CurrentPlayer) = Mid(Text, 8)
            ElseIf Mid(Text, 1, 6) = "##Type" Then
                PlayerType(CurrentPlayer) = Val(Mid(Text, 8))
            End If
        Next
    Next
    Close
    Imgplayer.Visible = False
    ScreenX = Int(BuildingX / 10)
    ScreenY = Int(BuildingY / 10)
    PlayerX = BuildingX - (ScreenX * 10)
    PlayerY = BuildingY - (ScreenY * 10)
    ScreenX = ScreenX + 1
    ScreenY = ScreenY + 1
    Imgplayer.Top = PlayerY * 375 - 240
    Imgplayer.Left = PlayerX * 375 - 135
    Anim = 1
    BuildingX = 0
    BuildingY = 0
    RenderMap
    Imgplayer.Visible = True
    TmrPlayer.Enabled = True
End Sub

Sub ClearData()
    For ClearPlayer = 1 To 800
        PlayerName(ClearPlayer) = ""
        CharX(ClearPlayer) = 0
        CharY(ClearPlayer) = 0
        PlayerType(ClearPlayer) = 0
        PlayerText(ClearPlayer) = ""
    Next
    For OuterLoop = 1 To 200
        For InnerLoop = 1 To 200
            MapPlayers(InnerLoop, OuterLoop) = 0
        Next
    Next
End Sub
