VERSION 5.00
Begin VB.Form frmAnime1 
   Caption         =   "Animated  Form  [ Medium Speed ]"
   ClientHeight    =   3165
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   5775
   Icon            =   "frmAnime1.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Unload"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Info >> The interesting thing is  that for a compiled exe the 'form' 'Orginates' from it's own Icon."
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Set the 'Animated Form' as the 'StartUp Object'  before compiling."
      Height          =   780
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You must exicute it after compiling to get the actual result."
      Height          =   540
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4920
   End
End
Attribute VB_Name = "frmAnime1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUnload_Click()
    Unload Me
End Sub

'Change the Animation Speed
Private Sub Form_Load()
    AnimateForm Me, aLoad, aMedium
End Sub

'Change the Animation Speed
Private Sub Form_Unload(Cancel As Integer)
    AnimateForm Me, aUnload, aMedium
End Sub

