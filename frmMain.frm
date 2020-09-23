VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6090
   ControlBox      =   0   'False
   FillColor       =   &H00808000&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5085
      TabIndex        =   47
      ToolTipText     =   "Minimize"
      Top             =   3570
      Width           =   345
   End
   Begin VB.CheckBox chkSound 
      Caption         =   "Sound"
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Top             =   5820
      Value           =   1  'Checked
      Width           =   885
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3645
      TabIndex        =   43
      ToolTipText     =   "Turn Sound Off"
      Top             =   3570
      Width           =   345
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   2220
      Top             =   4710
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3990
      TabIndex        =   25
      ToolTipText     =   "Start a New Game"
      Top             =   3570
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5430
      TabIndex        =   24
      ToolTipText     =   "Exit"
      Top             =   3570
      Width           =   345
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   23
      ToolTipText     =   "Help"
      Top             =   3570
      Width           =   345
   End
   Begin VB.PictureBox led4 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2070
      Picture         =   "frmMain.frx":000C
      ScaleHeight     =   225
      ScaleWidth      =   285
      TabIndex        =   14
      Top             =   3480
      Width           =   285
   End
   Begin VB.CommandButton cmdDecode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Decode"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      MaskColor       =   &H000000C0&
      TabIndex        =   13
      Tag             =   "0"
      Top             =   2940
      UseMaskColor    =   -1  'True
      Width           =   1125
   End
   Begin VB.CommandButton cmd0 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   690
      MaskColor       =   &H000000C0&
      TabIndex        =   12
      Top             =   2940
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   690
      MaskColor       =   &H000000C0&
      TabIndex        =   11
      Top             =   2460
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton cmd9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2010
      MaskColor       =   &H000000C0&
      TabIndex        =   10
      Top             =   2460
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton cmd8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      MaskColor       =   &H000000C0&
      TabIndex        =   9
      Top             =   2460
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2010
      MaskColor       =   &H000000C0&
      TabIndex        =   8
      Top             =   1950
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton cmd5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      MaskColor       =   &H000000C0&
      TabIndex        =   7
      Top             =   1950
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   690
      MaskColor       =   &H000000C0&
      TabIndex        =   6
      Top             =   1950
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2010
      MaskColor       =   &H000000C0&
      TabIndex        =   5
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      MaskColor       =   &H000000C0&
      TabIndex        =   4
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   690
      MaskColor       =   &H000000C0&
      TabIndex        =   3
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.PictureBox led2 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1215
      Picture         =   "frmMain.frx":0407
      ScaleHeight     =   225
      ScaleWidth      =   285
      TabIndex        =   2
      Top             =   3480
      Width           =   285
   End
   Begin VB.PictureBox led3 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1635
      Picture         =   "frmMain.frx":0802
      ScaleHeight     =   225
      ScaleWidth      =   285
      TabIndex        =   1
      Top             =   3480
      Width           =   285
   End
   Begin VB.PictureBox led1 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   780
      Picture         =   "frmMain.frx":0BFD
      ScaleHeight     =   225
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   3480
      Width           =   285
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2505
      Left            =   3210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   1020
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ALARM CODE"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   825
      TabIndex        =   48
      Top             =   390
      Width           =   1545
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   3240
      Top             =   3510
      Width           =   2625
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3735
      Left            =   240
      Top             =   240
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   3240
      X2              =   5820
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "CODE GREEN YELLOW"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   3240
      TabIndex        =   45
      Top             =   750
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "The variables below keep track of which digit box will recieve the next user chosen #"
      Height          =   645
      Left            =   3270
      TabIndex        =   42
      Top             =   4350
      Width           =   2565
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Pushed3"
      Height          =   255
      Left            =   4410
      TabIndex        =   41
      Top             =   5940
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Pushed2"
      Height          =   255
      Left            =   4410
      TabIndex        =   40
      Top             =   5550
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Pushed1"
      Height          =   255
      Left            =   4410
      TabIndex        =   39
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "lblGreen"
      Height          =   225
      Left            =   1050
      TabIndex        =   38
      Top             =   5190
      Width           =   1065
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "lblYellow"
      Height          =   225
      Left            =   1050
      TabIndex        =   37
      Top             =   5580
      Width           =   1125
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "lblTotalYellow"
      Height          =   225
      Left            =   1050
      TabIndex        =   36
      Top             =   5970
      Width           =   1005
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Texture"
      Height          =   225
      Left            =   2550
      TabIndex        =   35
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblTotalYellow 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   540
      TabIndex        =   34
      Top             =   5910
      Width           =   405
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   3735
      Left            =   3210
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DATA RECORDER"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   450
      Left            =   3240
      TabIndex        =   33
      Top             =   330
      Width           =   2595
   End
   Begin VB.Label lblGreen 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   540
      TabIndex        =   31
      Top             =   5130
      Width           =   405
   End
   Begin VB.Label lblYellow 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   540
      TabIndex        =   30
      Top             =   5520
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "secret code"
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   540
      TabIndex        =   29
      Top             =   4260
      Width           =   1500
   End
   Begin VB.Label pushed3 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3960
      TabIndex        =   28
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label pushed2 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3960
      TabIndex        =   27
      Top             =   5490
      Width           =   345
   End
   Begin VB.Label pushed1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3960
      TabIndex        =   26
      Top             =   5100
      Width           =   345
   End
   Begin VB.Label newcode4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1650
      TabIndex        =   22
      Top             =   4530
      Width           =   315
   End
   Begin VB.Label newcode3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1290
      TabIndex        =   21
      Top             =   4530
      Width           =   315
   End
   Begin VB.Label newcode2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   930
      TabIndex        =   20
      Top             =   4530
      Width           =   315
   End
   Begin VB.Label newcode1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   570
      TabIndex        =   19
      Top             =   4530
      Width           =   315
   End
   Begin VB.Label code4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   435
      Left            =   2085
      TabIndex        =   18
      Top             =   750
      Width           =   375
   End
   Begin VB.Label code3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   435
      Left            =   1635
      TabIndex        =   17
      Top             =   750
      Width           =   375
   End
   Begin VB.Label code2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   435
      Left            =   1170
      TabIndex        =   16
      Top             =   750
      Width           =   375
   End
   Begin VB.Label code1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   435
      Left            =   720
      TabIndex        =   15
      Top             =   750
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2130
      Picture         =   "frmMain.frx":0FF8
      Top             =   5370
      Width           =   390
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "zHUD"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   810
      Left            =   3240
      TabIndex        =   46
      Top             =   240
      Width           =   2595
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   240
      Top             =   240
      Width           =   2625
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub disable_all()

cmd1.Enabled = False
cmd2.Enabled = False
cmd3.Enabled = False
cmd4.Enabled = False
cmd5.Enabled = False
cmd6.Enabled = False
cmd7.Enabled = False
cmd8.Enabled = False
cmd9.Enabled = False
cmd0.Enabled = False

End Sub

Public Sub enable_all()


cmd1.Enabled = True
cmd2.Enabled = True
cmd3.Enabled = True
cmd4.Enabled = True
cmd5.Enabled = True
cmd6.Enabled = True
cmd7.Enabled = True
cmd8.Enabled = True
cmd9.Enabled = True
cmd0.Enabled = True

End Sub

Private Sub GetNewCode()

Dim Num1 As Integer
Dim Num2 As Integer
Dim Num3 As Integer
Dim Num4 As Integer

'RndNum is a function in the custom.bas which picks a random #
Num1 = RndNum(0, 9)

'The following lines make sure a digit isn't chosen twice
100
Num2 = RndNum(0, 9)
If Num2 = Num1 Then GoTo 100

200
Num3 = RndNum(0, 9)
If Num3 = Num1 Or Num3 = Num2 Then GoTo 200

300
Num4 = RndNum(0, 9)
If Num4 = Num1 Or Num4 = Num2 Or _
Num4 = Num3 Then GoTo 300

'Assign our new code to it's hidden location
newcode1.Caption = Num1
newcode2.Caption = Num2
newcode3.Caption = Num3
newcode4.Caption = Num4

End Sub

Public Sub Tone()

If chkSound.Value = 1 Then
PlayWav (App.Path + "\button.wav")

Else
End If

End Sub

Private Sub cmd0_Click()

If pushed1 = 0 Then
code1.Caption = 0
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 0
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 0
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 0
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub


Private Sub cmd1_Click()

'We check to see in which position to place this digit
If pushed1 = 0 Then
code1.Caption = 1
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 1
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 1
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 1
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub

Private Sub cmd2_Click()

If pushed1 = 0 Then
code1.Caption = 2
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 2
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 2
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 2
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub


Private Sub cmd3_Click()

If pushed1 = 0 Then
code1.Caption = 3
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 3
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 3
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 3
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub

Private Sub cmd4_Click()

If pushed1 = 0 Then
code1.Caption = 4
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 4
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 4
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 4
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub


Private Sub cmd5_Click()

If pushed1 = 0 Then
code1.Caption = 5
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 5
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 5
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 5
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub


Private Sub cmd6_Click()

If pushed1 = 0 Then
code1.Caption = 6
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 6
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 6
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 6
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub


Private Sub cmd7_Click()

If pushed1 = 0 Then
code1.Caption = 7
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 7
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 7
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 7
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub


Private Sub cmd8_Click()

If pushed1 = 0 Then
code1.Caption = 8
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 8
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 8
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 8
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub


Private Sub cmd9_Click()

If pushed1 = 0 Then
code1.Caption = 9
pushed1 = 1

ElseIf pushed2 = 0 Then
code2.Caption = 9
pushed2 = 1

ElseIf pushed3 = 0 Then
code3.Caption = 9
pushed3 = 1

Else
cmdDecode.Enabled = True
code4.Caption = 9
pushed1 = 0
pushed2 = 0
pushed3 = 0

End If

Call Tone

End Sub



Private Sub cmdDecode_Click()

'Use the decode button tag to keep track of the number of times
'it's been pressed.
cmdDecode.Tag = cmdDecode.Tag + 1
cmdDecode.Enabled = False

'Reset all lights to red
led1.Picture = LoadPicture(App.Path & "\red.gif")
led2.Picture = LoadPicture(App.Path & "\red.gif")
led3.Picture = LoadPicture(App.Path & "\red.gif")
led4.Picture = LoadPicture(App.Path & "\red.gif")

'Pushed captions keep track of which box the next digit will be
'entered into
pushed1.Caption = 0
pushed2.Caption = 0
pushed3.Caption = 0

'Variables for checking for correct #'s and Positions
Dim chkNum1 As Integer
Dim chkNum2 As Integer
Dim chkNum3 As Integer
Dim chkNum4 As Integer

Dim chkPos1 As Integer
Dim chkPos2 As Integer
Dim chkPos3 As Integer
Dim chkPos4 As Integer

Dim TotalGreen As Integer
Dim TotalYellow As Integer

' Check for correct access code
If code1.Caption = newcode1.Caption _
And code2.Caption = newcode2.Caption _
And code3.Caption = newcode3.Caption _
And code4.Caption = newcode4.Caption Then

'Winner (code has been cracked)
led1.Picture = LoadPicture(App.Path & "\green.gif")
led2.Picture = LoadPicture(App.Path & "\green.gif")
led3.Picture = LoadPicture(App.Path & "\green.gif")
led4.Picture = LoadPicture(App.Path & "\green.gif")

'Play sound if enabled
If chkSound.Value = 1 Then
PlayWav (App.Path + "\granted.wav")
Else
End If

'Disables all buttons until a new game is started
Call disable_all

'Text displayed in the data recorder window
'VbCrLf is a carriage return, line feed (the same as pressing
'Enter to start a new line
txtData.Text = ""
txtData.Text = vbCrLf + vbCrLf + vbCrLf + vbCrLf + _
"   ACCESS GRANTED"
txtData.Text = txtData.Text + vbCrLf + vbCrLf + _
"     " + cmdDecode.Tag + " ATTEMPTS"

Exit Sub

   Else
End If

'Code is incorrect so we continue....
'Check for existence of the first numer
If code1.Caption = newcode1.Caption _
Or code1.Caption = newcode2.Caption _
Or code1.Caption = newcode3.Caption _
Or code1.Caption = newcode4.Caption Then

'1 means yes a correct number is present
chkNum1 = 1
End If

'then check for its position
If code1.Caption = newcode1.Caption Then
'1 means yes a number is in the correct postion
chkPos1 = 1
End If

'Check for existence of second number
If code2.Caption = newcode1.Caption _
Or code2.Caption = newcode2.Caption _
Or code2.Caption = newcode3.Caption _
Or code2.Caption = newcode4.Caption Then
chkNum2 = 1
End If

'then check for its position
If code2.Caption = newcode2.Caption Then
chkPos2 = 1
End If


'Check for existence of third number
If code3.Caption = newcode1.Caption _
Or code3.Caption = newcode2.Caption _
Or code3.Caption = newcode3.Caption _
Or code3.Caption = newcode4.Caption Then
chkNum3 = 1
End If

'then check for its position
If code3.Caption = newcode3.Caption Then
chkPos3 = 1
End If


'Check for existence of fourth number
If code4.Caption = newcode1.Caption _
Or code4.Caption = newcode2.Caption _
Or code4.Caption = newcode3.Caption _
Or code4.Caption = newcode4.Caption Then
chkNum4 = 1
End If

'then check for its position
If code4.Caption = newcode4.Caption Then
chkPos4 = 1
End If


'Check for the sum of yellow conditions
If chkNum1 = 1 Then
TotalYellow = TotalYellow + 1
Else
End If

If chkNum2 = 1 Then
TotalYellow = TotalYellow + 1
Else
End If

If chkNum3 = 1 Then
TotalYellow = TotalYellow + 1
Else
End If

If chkNum4 = 1 Then
TotalYellow = TotalYellow + 1
Else
End If

lblYellow.Caption = TotalYellow


'Check for the sum of green conditions
If chkPos1 = 1 Then
TotalGreen = TotalGreen + 1
Else
End If

If chkPos2 = 1 Then
TotalGreen = TotalGreen + 1
Else
End If

If chkPos3 = 1 Then
TotalGreen = TotalGreen + 1
Else
End If

If chkPos4 = 1 Then
TotalGreen = TotalGreen + 1
Else
End If

lblGreen.Caption = TotalGreen


'Assign green lights to panel
If TotalGreen = 1 Then
led1.Picture = LoadPicture(App.Path & "\green.gif")

ElseIf TotalGreen = 2 Then
led1.Picture = LoadPicture(App.Path & "\green.gif")
led2.Picture = LoadPicture(App.Path & "\green.gif")

ElseIf TotalGreen = 3 Then
led1.Picture = LoadPicture(App.Path & "\green.gif")
led2.Picture = LoadPicture(App.Path & "\green.gif")
led3.Picture = LoadPicture(App.Path & "\green.gif")

Else
End If

'Filter out the yellow lights that correspond with green ones.
'Then assign yellow lights to panel
If TotalGreen = 0 Then

 If TotalYellow = 1 Then
 led1.Picture = LoadPicture(App.Path & "\yellow.gif")
 
 ElseIf TotalYellow = 2 Then
 led1.Picture = LoadPicture(App.Path & "\yellow.gif")
 led2.Picture = LoadPicture(App.Path & "\yellow.gif")
 
 ElseIf TotalYellow = 3 Then
 led1.Picture = LoadPicture(App.Path & "\yellow.gif")
 led2.Picture = LoadPicture(App.Path & "\yellow.gif")
 led3.Picture = LoadPicture(App.Path & "\yellow.gif")
 
 ElseIf TotalYellow = 4 Then
 led1.Picture = LoadPicture(App.Path & "\yellow.gif")
 led2.Picture = LoadPicture(App.Path & "\yellow.gif")
 led3.Picture = LoadPicture(App.Path & "\yellow.gif")
 led4.Picture = LoadPicture(App.Path & "\yellow.gif")
 
 Else
 End If
 
Else
End If


If TotalGreen = 1 Then

 If TotalYellow = 2 Then
 led2.Picture = LoadPicture(App.Path & "\yellow.gif")
 
 ElseIf TotalYellow = 3 Then
 led2.Picture = LoadPicture(App.Path & "\yellow.gif")
 led3.Picture = LoadPicture(App.Path & "\yellow.gif")
 
 ElseIf TotalYellow = 4 Then
 led2.Picture = LoadPicture(App.Path & "\yellow.gif")
 led3.Picture = LoadPicture(App.Path & "\yellow.gif")
 led4.Picture = LoadPicture(App.Path & "\yellow.gif")
 
 Else
 End If
 
Else
End If

If TotalGreen = 2 Then

 If TotalYellow = 3 Then
 led3.Picture = LoadPicture(App.Path & "\yellow.gif")
 
 ElseIf TotalYellow = 4 Then
 led3.Picture = LoadPicture(App.Path & "\yellow.gif")
 led4.Picture = LoadPicture(App.Path & "\yellow.gif")
 
 Else
 End If
 
Else
End If

'Final filter to eliminate extra yellows for easier counting
lblTotalYellow.Caption = TotalYellow - TotalGreen

'Writing code status to the data recorder.
txtData.Text = txtData.Text + " " + code1.Caption _
+ code2.Caption + code3.Caption + code4.Caption _
+ "    " + lblGreen.Caption + "      " + lblTotalYellow.Caption _
+ vbCrLf

'Play sound if enabled
If chkSound.Value = 1 Then
PlayWav (App.Path + "\denied.wav")
Else
End If

'Check to see if the last attempt has been entered.
If cmdDecode.Tag = 15 Then

'Timer that enables the alarm audio loop.
Timer1.Enabled = True

'Reveal the real alarm code to the player
code1.Caption = newcode1.Caption
code2.Caption = newcode2.Caption
code3.Caption = newcode3.Caption
code4.Caption = newcode4.Caption

Call disable_all

'YOU LOSE....Text written to data recorder
txtData.Text = ""
txtData.Text = vbCrLf + vbCrLf + vbCrLf + vbCrLf + "   ALARM ACTIVATED"
txtData.Text = txtData.Text + vbCrLf + vbCrLf + "    DECODE FAILED"

Exit Sub

Else
End If

'Clear the code for the next attempt.
code1.Caption = "-"
code2.Caption = "-"
code3.Caption = "-"
code4.Caption = "-"

End Sub


Private Sub cmdExit_Click()

End

End Sub


Private Sub cmdHelp_Click()

frmHelp.Visible = True

End Sub

Private Sub cmdMinimize_Click()

'Minimize
frmMain.WindowState = 1

End Sub

Private Sub cmdNewGame_Click()

'Enable all buttons
Call enable_all

'Stop the alarm
Timer1.Enabled = False
cmdDecode.Tag = 0

'Reset all data to begin a new game
pushed1.Caption = 0
pushed2.Caption = 0
pushed3.Caption = 0

code1.Caption = "-"
code2.Caption = "-"
code3.Caption = "-"
code4.Caption = "-"

led1.Picture = LoadPicture(App.Path & "\red.gif")
led2.Picture = LoadPicture(App.Path & "\red.gif")
led3.Picture = LoadPicture(App.Path & "\red.gif")
led4.Picture = LoadPicture(App.Path & "\red.gif")

'Subroutine which picks a new random code.
Call GetNewCode

'Clear the data recorder
txtData.Text = ""

'Play sound if enabled
If chkSound.Value = 1 Then
PlayWav (App.Path + "\restart.wav")
Else
End If

End Sub

Private Sub cmdSound_Click()

'This is a good example of how to toggle a button for
'on/off type uses

If chkSound.Value = 1 Then
chkSound.Value = 0
cmdSound.ToolTipText = "Turn Sound On"

Else
chkSound.Value = 1
cmdSound.ToolTipText = "Turn Sound Off"

End If

End Sub

Private Sub Form_Load()

Call disable_all
cmdDecode.Enabled = False

txtData.Text = vbCrLf + vbCrLf + "  CODE CRACKER 2001" + vbCrLf _
+ vbCrLf + "    By Roger Lee"

End Sub

Private Sub Form_Paint()
Dim intX As Integer
Dim intY As Integer

For intX = 0 To Me.Width Step Image1.Width
    For intY = 0 To Me.Height Step Image1.Height
        PaintPicture Image1, intX, intY
    Next intY
Next intX
End Sub








Private Sub Timer1_Timer()

'Play alarm sound if enabled
If chkSound.Value = 1 Then
PlayWav (App.Path + "\alarm.wav")
Else
End If

End Sub


