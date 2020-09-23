VERSION 5.00
Begin VB.Form frmAnime2 
   Caption         =   "Animated  Form  [ Medium Speed ]"
   ClientHeight    =   3045
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   5955
   FillColor       =   &H00FF0000&
   Icon            =   "frmAnime2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3045
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Unload"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Assumption >> Probably ' Form.Fillcolor'  has a little significance"
      Height          =   480
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   4740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAnime2.frx":08CA
      Height          =   780
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   4920
   End
End
Attribute VB_Name = "frmAnime2"
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

