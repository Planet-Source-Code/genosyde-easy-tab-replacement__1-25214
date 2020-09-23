VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Replacement for Tabs"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   4
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4455
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label4 
         Caption         =   $"frmMain.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "Black Data"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Created by:  "
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Replacement for Tabs"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   3
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4455
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label5 
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   2
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4455
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label6 
         Caption         =   "Tools"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   1
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4455
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Label7 
         Caption         =   "View"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   0
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4455
      TabIndex        =   6
      Top             =   720
      Width           =   4455
      Begin VB.Label Label8 
         Caption         =   "Main"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picBorder 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Label lblMenuItem 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "About"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblMenuItem 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblMenuItem 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Tools"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblMenuItem 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "View"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblMenuItem 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Main"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      X1              =   120
      X2              =   4560
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      X1              =   4560
      X2              =   4560
      Y1              =   720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      BorderWidth     =   3
      X1              =   120
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   720
      Y2              =   3720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' Forces VB to recognize all variables as a 'type'
Dim intX As Integer ' Set variable (intX) as a Number

Private Sub lblMenuItem_Click(Index As Integer)
For intX = 0 To 4 ' Set variable (intX) to loops through numbers 0-4
    lblMenuItem(intX).FontBold = False ' Take of the Bold on all Menu Items
    lblMenuItem(Index).FontBold = True ' Make Bold the Menu Item selected (Index)
    picFrame(intX).Visible = False ' Make all Frames disappear
    picFrame(Index).Visible = True ' Make Frame selected by Menu Item (Index) appear
Next intX ' Loop
End Sub

'EASY

'I was going to Offically make this an OCX, but I was to lazy.
'If you decide to make this an ActiveX Control, just give me some credit.
'... also I will help as well.
