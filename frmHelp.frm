VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   2910
      TabIndex        =   12
      Top             =   6090
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   510
      Picture         =   "frmHelp.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   285
      TabIndex        =   2
      Top             =   5490
      Width           =   285
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   510
      Picture         =   "frmHelp.frx":03FB
      ScaleHeight     =   225
      ScaleWidth      =   285
      TabIndex        =   1
      Top             =   4290
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   510
      Picture         =   "frmHelp.frx":07F6
      ScaleHeight     =   225
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   4890
      Width           =   285
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Lights only indicate overall code accuracy. They do not indicate which digits are valid and which are not."
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   360
      TabIndex        =   13
      Top             =   3540
      Width           =   5535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Indicates that your code has a digit that is not present in the access code."
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   960
      TabIndex        =   11
      Top             =   5460
      Width           =   5115
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Indicates that your code has a valid digit, but in the incorrect location."
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   990
      TabIndex        =   10
      Top             =   4860
      Width           =   5265
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Indicates that your code has a valid digit in the correct location."
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   990
      TabIndex        =   9
      Top             =   4260
      Width           =   4965
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LIGHT LEGEND"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESS"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   330
      TabIndex        =   7
      Top             =   1350
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GOAL"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   390
      Width           =   705
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":0BF1
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   5625
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "1) Enter your code and press the decode button. Codes will never have two of the same numbers."
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   630
      Left            =   330
      TabIndex        =   4
      Top             =   1680
      Width           =   5595
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   6405
      Left            =   210
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "You must determine the alarm code in 15 attempts or less. If you fail the alarm will sound. "
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   360
      TabIndex        =   3
      Top             =   660
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Picture         =   "frmHelp.frx":0C9E
      Top             =   0
      Width           =   390
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   6345
      Left            =   240
      Top             =   270
      Width           =   5865
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

Unload Me

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


