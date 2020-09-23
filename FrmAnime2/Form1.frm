VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Professional Animation - Try At Different Controlls and Speed"
   ClientHeight    =   5355
   ClientLeft      =   75
   ClientTop       =   705
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad2 
      Caption         =   "Load"
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   10
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad2 
      Caption         =   "Load"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Index           =   9
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Index           =   8
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Index           =   7
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Index           =   6
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Index           =   5
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Index           =   4
      Left            =   6120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Try different Speeds manually  ......."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   11
      Top             =   2280
      Width           =   4650
   End
   Begin VB.Menu mnufrm 
      Caption         =   "Form"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu tyt 
         Caption         =   ""
      End
      Begin VB.Menu gf 
         Caption         =   ""
      End
      Begin VB.Menu trytry 
         Caption         =   ""
      End
      Begin VB.Menu yryt 
         Caption         =   ""
      End
      Begin VB.Menu fdd 
         Caption         =   ""
      End
      Begin VB.Menu mnuLoad1 
         Caption         =   "Load"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click(Index As Integer)
    frmAnime1.Show
End Sub

Private Sub cmdLoad2_Click(Index As Integer)
    frmAnime2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Integer
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Please tell Me") = vbYes Then
        X = MsgBox("(  PLEASE 'RATE' THIS CODE  ).I want to know how do you rate this code.Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "ThankYou")
    Else
        X = MsgBox("( PLEASE GIVE FEEDBACK ) to improve this code.Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "Please Give FeedBack")
    End If
    If X = vbOK Then Clipboard.SetText ("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57829&lngWId=1")
End Sub

Private Sub mnuLoad_Click()
    frmAnime2.Show
End Sub

Private Sub mnuLoad1_Click()
    frmAnime1.Show
End Sub
